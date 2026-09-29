"""
Delivery Operations - Order Cleaner & Auto-Router
-------------------------------------------------
Run:  streamlit run app.py
Optional admin lock:  set env var ADMIN_PASSWORD (or .streamlit/secrets.toml).

Pipeline
  1. Upload Excel -> map columns -> trim/clean text.
  2. Normalise phones (strip spaces/dashes/+60) and addresses (Jln->Jalan, etc.).
  3. Flag duplicates: repeated order numbers, repeat phones, phones with different
     addresses, and fuzzy-similar addresses (RapidFuzz) with different/same phones.
  4. Cluster orders by postcode / area name, auto-assign to riders.
  5. Enforce per-rider caps; overflow goes to the nearest rider with spare capacity.
"""
import hmac
import io
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from itertools import combinations
from statistics import median

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

st.set_page_config(page_title="Delivery Order Router", page_icon="🚚", layout="wide")

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #
DEFAULT_CAPS = {"arif": 15, "fairuz": 14, "shah": 18}
FALLBACK_CAP = 15
NEW_RIDER_DIST = 1500  # "distance" for a rider with no orders yet (postcode units)
OVERFLOW = "Overflow → backup"

ABBREV = {
    "jln": "jalan", "lrg": "lorong", "tmn": "taman", "kg": "kampung", "bdr": "bandar",
    "psn": "persiaran", "lbh": "lebuh", "sek": "seksyen", "apt": "apartment",
    "blk": "blok", "bkt": "bukit", "bt": "batu", "sg": "sungai", "seri": "sri",
    "pj": "petaling jaya", "kl": "kuala lumpur", "jb": "johor bahru", "pjs": "petaling jaya selatan",
}
STATES = {
    "selangor", "kuala lumpur", "wilayah persekutuan", "putrajaya", "johor", "penang",
    "pulau pinang", "perak", "kedah", "kelantan", "terengganu", "pahang", "negeri sembilan",
    "melaka", "malacca", "sabah", "sarawak", "perlis", "labuan", "malaysia",
    "wilayah persekutuan kuala lumpur", "federal territory of kuala lumpur", "wp kuala lumpur",
}
POSTCODE_RE = re.compile(r"(?<!\d)(\d{5})(?!\d)")
PHONE_RE = re.compile(r"(?<!\d)(?:\+?6?0)\d(?:[\s\-]?\d){7,9}(?!\d)")
UNITLIKE_RE = re.compile(r"^[a-z]{0,2}-?\d", re.I)
SEV_RANK = {"Low": 1, "Medium": 2, "High": 3}
LEVEL_FILL = {"High": "F8CBAD", "Medium": "FFE699", "Low": "FFF2CC"}


# --------------------------------------------------------------------------- #
# Text helpers
# --------------------------------------------------------------------------- #
def normalize_address(addr: str) -> str:
    s = re.sub(r"\bk\.\s?l\b\.?", "kl", str(addr).lower())
    s = re.sub(r"[.,;#()]", " ", s)
    s = re.sub(r"\bno\s*(?=\d)", "no ", s)
    return " ".join(ABBREV.get(t, t) for t in s.split() if t != "malaysia")


def unit_key(norm: str) -> str:
    """House / unit number at the start of an address (e.g. '12', 'a-12-3')."""
    m = re.match(r"^(?:(?:no|lot|unit|blok|apartment)\s+)?([a-z]{0,2}-?\d+[a-z]?(?:[-/]\d+[a-z]?)*)\b", norm)
    return m.group(1) if m else ""


def normalize_phone(raw: str) -> str:
    d = re.sub(r"\D", "", raw)
    if d.startswith("600"):
        d = d[2:]
    elif d.startswith("60"):
        d = "0" + d[2:]
    elif not d.startswith("0"):
        d = "0" + d
    return d if 9 <= len(d) <= 11 else ""


def extract_phones(text: str) -> list:
    found = []
    for m in PHONE_RE.finditer(str(text)):
        p = normalize_phone(m.group(0))
        if p and p not in found:
            found.append(p)
    return found


def extract_postcode_area(addr: str):
    """Return (postcode, area name) parsed from a Malaysian-style address."""
    pc, area = "", ""
    m = POSTCODE_RE.search(addr)
    if m:
        pc = m.group(1)
        tail = addr[m.end():].strip(" ,-")
        cand = re.split(r"[,\n]", tail)[0].strip()
        cand = re.sub(r"\s+(?:%s)$" % "|".join(sorted(STATES, key=len, reverse=True)), "", cand, flags=re.I)
        if cand and cand.lower() not in STATES and not POSTCODE_RE.fullmatch(cand):
            area = cand
    if not area:
        parts = [p.strip() for p in re.split(r"[,\n]", addr) if p.strip()]
        parts = [p for p in parts if not POSTCODE_RE.fullmatch(p) and p.lower() not in STATES
                 and not UNITLIKE_RE.match(p)]
        area = POSTCODE_RE.sub("", parts[-1]).strip() if parts else ""
    return pc, area.title()


def tidy_address(addr: str) -> str:
    """Drop empty / repeated comma segments (e.g. '40170, 40170, Shah Alam')."""
    seen, seen_pc, out = set(), set(), []
    for seg in (x.strip() for x in addr.split(",")):
        if not seg or seg.lower() in seen or (POSTCODE_RE.fullmatch(seg) and seg in seen_pc):
            continue
        seen.add(seg.lower())
        seen_pc.update(POSTCODE_RE.findall(seg))
        out.append(seg)
    return ", ".join(out)


def parse_order_blob(text) -> dict:
    """Split a pasted order cell (name / phone(s) / address / notes) into its parts.
    Also works for a plain one-line address."""
    lines = [l.strip() for l in re.split(r"[\r\n]+", str(text)) if l.strip()]
    has_phone = lambda l: bool(PHONE_RE.search(l))
    phones = []
    for l in lines:
        phones += [p for p in extract_phones(l) if p not in phones]
    customer = re.sub(r"\s+", " ", PHONE_RE.sub("", lines[0])).strip(" ,-()") if lines else ""

    start = next((k for k, l in enumerate(lines) if POSTCODE_RE.search(l) and not has_phone(l)), None)
    if start is None:  # no postcode: first plausible address line after the name line
        rest = [k for k in range(1, len(lines)) if not has_phone(lines[k])]
        start = next((k for k in rest if "," in lines[k]), None)
        if start is None:
            start = next((k for k in rest if re.search(r"\d", lines[k])), None)
        if start is None and lines and not has_phone(lines[0]) and ("," in lines[0] or len(lines) == 1):
            start = 0
    idx = []
    if start is not None:
        idx = [start]
        while lines[idx[-1]].endswith(",") and idx[-1] + 1 < len(lines) and len(idx) < 3:
            idx.append(idx[-1] + 1)
    notes = [l for k, l in enumerate(lines) if k not in idx and k != 0 and not has_phone(l)]
    return dict(customer=customer, phones=phones, notes=" | ".join(notes),
                address=tidy_address(" ".join(lines[k] for k in idx)))


def geo_dist(c1, a1, c2, a2) -> float:
    """Distance proxy: numeric postcode gap; falls back to area-name similarity."""
    if c1 is not None and c2 is not None:
        return abs(c1 - c2)
    if a1 and a2:
        if a1 == a2 or a1 in a2 or a2 in a1:
            return 0
        return 3000 + (100 - fuzz.ratio(a1, a2)) * 10
    return 9000


# --------------------------------------------------------------------------- #
# Duplicate flagging
# --------------------------------------------------------------------------- #
def flag_duplicates(order_nums, addr_norms, units, phones, threshold, ignore_diff_unit, refs):
    n = len(order_nums)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    flags = [[] for _ in range(n)]

    def add(i, sev, text):
        flags[i].append((sev, text))

    def label(i):
        return refs[i]

    # 1) duplicate order numbers
    by_order = defaultdict(list)
    for i, o in enumerate(order_nums):
        if o:
            by_order[o.lower()].append(i)
    for ids in by_order.values():
        if len(ids) > 1:
            for i in ids:
                add(i, "High", f"Duplicate order number ({len(ids)} rows)")
            for i in ids[1:]:
                union(ids[0], i)

    # 2) same phone on several orders
    by_phone = defaultdict(list)
    for i, ps in enumerate(phones):
        for p in ps:
            by_phone[p].append(i)
    for p, ids in by_phone.items():
        if len(ids) > 1:
            for i in ids:
                add(i, "Low", f"Phone {p} on {len(ids)} orders")
            for i in ids[1:]:
                union(ids[0], i)
            if len(ids) <= 25 and any(
                fuzz.token_sort_ratio(addr_norms[a], addr_norms[b]) < threshold
                for a, b in combinations(ids, 2)
            ):
                for i in ids:
                    add(i, "Medium", f"Same phone {p}, different addresses")

    # 3) fuzzy-similar addresses (chunked to keep memory small)
    valid = [i for i in range(n) if addr_norms[i]]
    texts = [addr_norms[i] for i in valid]
    CH = 1000
    for s in range(0, len(texts), CH):
        block = process.cdist(texts[s:s + CH], texts, scorer=fuzz.token_sort_ratio,
                              dtype=np.uint8, score_cutoff=threshold, workers=-1)
        rows, cols = np.nonzero(block)
        for r, c in zip(rows + s, cols):
            if c <= r:
                continue
            i, j = valid[r], valid[c]
            if ignore_diff_unit and units[i] and units[j] and units[i] != units[j]:
                continue
            score = int(block[r - s, c])
            pi, pj = set(phones[i]), set(phones[j])
            if pi and pj:
                sev, why = ("High", "same phone") if pi & pj else ("Medium", "different phone")
            else:
                sev, why = "Low", "no phone to compare"
            add(i, sev, f"Similar address to {label(j)} ({score}%, {why})")
            add(j, sev, f"Similar address to {label(i)} ({score}%, {why})")
            union(i, j)

    # finalise
    comps = defaultdict(list)
    for i in range(n):
        if flags[i]:
            comps[find(i)].append(i)
    gid = {}
    for k, members in enumerate(sorted(comps.values(), key=lambda m: m[0]), 1):
        for m in members:
            gid[m] = f"G{k:03d}"

    levels, reasons, groups = [], [], []
    for i in range(n):
        if not flags[i]:
            levels.append(""); reasons.append(""); groups.append("")
            continue
        texts_i = list(dict.fromkeys(t for _, t in flags[i]))
        extra = f" | +{len(texts_i) - 4} more" if len(texts_i) > 4 else ""
        levels.append(max((s for s, _ in flags[i]), key=SEV_RANK.get))
        reasons.append(" | ".join(texts_i[:4]) + extra)
        groups.append(gid[i])
    return levels, reasons, groups


# --------------------------------------------------------------------------- #
# Assignment engine (clusters + capacity caps + overflow)
# --------------------------------------------------------------------------- #
@dataclass
class Rider:
    name: str
    cap: int
    available: bool
    backup_only: bool
    anchors: set = field(default_factory=set)  # (postcode int|None, area key)
    load: int = 0

    @property
    def free(self) -> int:
        return self.cap - self.load


def _flag(v, default):
    return default if pd.isna(v) else bool(v)


def parse_home(txt) -> set:
    out = set()
    for tok in re.split(r"[,;]", "" if pd.isna(txt) else str(txt)):
        tok = tok.strip()
        if not tok:
            continue
        out.add((int(tok), "") if POSTCODE_RE.fullmatch(tok) else (None, normalize_address(tok)))
    return out


def assign_orders(recs, roster, respect_existing):
    riders = {}
    for _, row in roster.iterrows():
        name = str(row.get("Rider", "") if not pd.isna(row.get("Rider")) else "").strip()
        if not name or name.lower() in {k.lower() for k in riders}:
            continue
        cap = FALLBACK_CAP if pd.isna(row["Max Parcels"]) else max(int(row["Max Parcels"]), 0)
        riders[name] = Rider(name, cap, _flag(row["Available"], True),
                             _flag(row["Backup Only"], False), parse_home(row.get("Home Areas")))
    active = {k: r for k, r in riders.items() if r.available}
    lookup = {k.lower(): k for k in active}

    n = len(recs)
    out, types, notes = [""] * n, [""] * n, [""] * n

    def take(i, r, kind):
        r.load += 1
        r.anchors.add((recs[i]["code"], recs[i]["area"]))
        out[i], types[i] = r.name, kind

    def rider_dist(r, code, area):
        if not r.anchors:
            return NEW_RIDER_DIST
        return min(geo_dist(c, a, code, area) for c, a in r.anchors)

    def pick(code, area, regular_only):
        cands = [r for r in active.values() if r.free > 0 and not (regular_only and r.backup_only)]
        if not cands:
            return None
        return min(cands, key=lambda r: (rider_dist(r, code, area), -r.free, not r.backup_only, r.name))

    # Step A: honour existing assignments up to each rider's cap
    pool = []
    for i, rec in enumerate(recs):
        orig = rec["orig"].strip()
        if respect_existing and orig:
            nm = lookup.get(orig.lower())
            if nm and active[nm].free > 0:
                take(i, active[nm], "Pre-assigned")
                continue
            notes[i] = f"{orig} at cap ({active[nm].cap})" if nm else f"{orig} unavailable / not on roster"
        pool.append(i)

    # Step B: cluster the rest (biggest clusters first) and hand them to the nearest rider
    clusters = defaultdict(list)
    for i in pool:
        clusters[recs[i]["cluster"]].append(i)
    for idxs in sorted(clusters.values(), key=len, reverse=True):
        code, area = recs[idxs[0]]["code"], recs[idxs[0]]["area"]
        primary = pick(code, area, regular_only=True)
        for i in idxs:
            if not notes[i] and primary and primary.free > 0:
                take(i, primary, "Auto (cluster)")
                continue
            r = pick(code, area, regular_only=False)  # nearest rider with spare capacity
            if r is None:
                types[i], notes[i] = "Unassigned", "All riders at capacity"
                continue
            if not notes[i]:
                notes[i] = f"{primary.name} at cap ({primary.cap})" if primary else "All regular riders at cap"
            take(i, r, OVERFLOW)

    over = Counter(o for o, t in zip(out, types) if t == OVERFLOW)
    loads = pd.DataFrame([{
        "Rider": r.name, "Max Parcels": r.cap, "Assigned": r.load,
        "Free Slots": r.cap - r.load if r.available else 0,
        "Utilization %": round(100 * r.load / r.cap, 1) if r.cap else 0.0,
        "Overflow Received": over.get(r.name, 0),
        "Available": r.available, "Backup Only": r.backup_only,
    } for r in riders.values()])
    return out, types, notes, loads


# --------------------------------------------------------------------------- #
# Pipeline + Excel export
# --------------------------------------------------------------------------- #
def run_pipeline(raw, cols, roster, threshold, ignore_unit, respect, first_row=2):
    order_c, addr_c, remark_c, rider_c, phone_c = cols
    df = raw.copy()
    df.insert(0, "Source Row", range(first_row, first_row + len(df)))  # row number in the uploaded sheet
    keep = df[addr_c].str.strip().ne("")
    if order_c:
        keep |= df[order_c].str.strip().ne("")
    df = df[keep].reset_index(drop=True)
    used = {order_c, addr_c, remark_c, rider_c, phone_c}
    empty = [c for c in df.columns if c not in used and df[c].astype(str).str.strip().eq("").all()]
    df = df.drop(columns=empty)

    # split each address cell into customer / phones / address / notes
    blobs = [parse_order_blob(t) for t in df[addr_c]]
    addr_clean = [b["address"] for b in blobs]
    norms = [normalize_address(a) for a in addr_clean]
    units = [unit_key(x) for x in norms]
    phones = []
    for i, b in enumerate(blobs):
        found = list(b["phones"])
        for c in (phone_c, remark_c):
            if c:
                found += [p for p in extract_phones(df[c].iat[i]) if p not in found]
        phones.append(found)

    # postcode / area; rows without a postcode borrow one from a known area name in the same file
    parsed = [extract_postcode_area(a) for a in addr_clean]
    seen = defaultdict(list)
    for pc, area in parsed:
        if pc and area:
            seen[normalize_address(area)].append(int(pc))
    area_code = {k: int(median(v)) for k, v in seen.items()}
    riders_in = df[rider_c] if rider_c else pd.Series([""] * len(df))
    recs, disp_pc = [], []
    for i, (pc, area) in enumerate(parsed):
        akey = normalize_address(area)
        code = int(pc) if pc else area_code.get(akey)
        if not pc and code is None:
            code = next((area_code[k] for k in sorted(area_code, key=len, reverse=True)
                         if len(k) >= 6 and k in norms[i]), None)
        disp_pc.append(pc or (f"~{code}" if code else ""))
        recs.append(dict(orig=riders_in.iat[i], code=code, area=akey,
                         cluster=str(code) if code else (f"area:{akey}" if akey else "unknown")))

    # order references: only trust values that look like real order IDs
    ids = []
    for i in range(len(df)):
        v = df[order_c].iat[i].strip() if order_c else ""
        ok = bool(re.search(r"\d", v)) and v.lower() != riders_in.iat[i].strip().lower()
        ids.append(v if ok else "")
    refs = [v or f"Row {r}" for v, r in zip(ids, df["Source Row"])]

    levels, reasons, groups = flag_duplicates(ids, norms, units, phones, threshold, ignore_unit, refs)
    riders_out, types, notes, loads = assign_orders(recs, roster, respect)

    rc = rider_c or "Rider"
    out = df.copy()
    if rider_c:
        out["Original Rider"] = df[rider_c]
    out[rc] = [r or "UNASSIGNED" for r in riders_out]
    out["Assignment Type"], out["Assignment Note"] = types, notes
    out["Order Ref"] = refs
    out["Customer"] = [b["customer"] for b in blobs]
    out["Parsed Address"] = addr_clean
    out["Delivery Notes"] = [b["notes"] for b in blobs]
    out["Postcode"], out["Area"] = disp_pc, [normalize_address(a).title() for _, a in parsed]
    out["Phone (Normalized)"] = [", ".join(p) for p in phones]
    out["Flag Level"], out["Flag Reason"], out["Dup Group"] = levels, reasons, groups

    out = out.sort_values([rc, "Postcode", "Area"], kind="stable",
                          key=lambda s: s.str.lower().replace("unassigned", "~~~")).reset_index(drop=True)
    fcols = ["Dup Group", "Flag Level", "Flag Reason", "Source Row", "Order Ref", "Customer",
             "Parsed Address", "Phone (Normalized)", rc, "Assignment Type"]
    flagged = out.loc[out["Flag Level"] != "", fcols].sort_values(["Dup Group", "Source Row"])
    return dict(routing=out, flagged=flagged, loads=loads, rider_col=rc,
                order_ids=(sum(bool(v) for v in ids), len(ids)))


def _style_sheet(ws):
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, col in enumerate(ws.columns, 1):
        w = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 10), 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def build_excel(res) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        res["routing"].to_excel(xw, sheet_name="Routing Sheet", index=False)
        res["flagged"].to_excel(xw, sheet_name="Flagged Duplicates", index=False)
        res["loads"].to_excel(xw, sheet_name="Rider Loads", index=False)
        for ws in xw.book.worksheets:
            _style_sheet(ws)
        ws = xw.book["Routing Sheet"]
        hdr = [c.value for c in ws[1]]
        fl, at = hdr.index("Flag Level"), hdr.index("Assignment Type")
        for row in ws.iter_rows(min_row=2):
            lvl = row[fl].value
            if lvl in LEVEL_FILL:
                for c in row:
                    c.fill = PatternFill("solid", fgColor=LEVEL_FILL[lvl])
            if row[at].value == OVERFLOW:
                row[at].fill = PatternFill("solid", fgColor="F4B084")
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# UI helpers
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def read_sheet(data: bytes, sheet: str, header: bool = True) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(data), sheet_name=sheet, dtype=str, header=0 if header else None).fillna("")
    df.columns = [f"Col {get_column_letter(i + 1)}" if (not header or str(c).startswith("Unnamed"))
                  else str(c).strip() for i, c in enumerate(df.columns)]
    return df


def guess(cols, keys, default=None):
    for k in keys:
        for c in cols:
            if k in str(c).lower():
                return c
    return default


def default_roster(raw, rider_col) -> pd.DataFrame:
    """Riders found in the file (case-insensitive). Caps: your known caps, else today's count in the file."""
    variants = defaultdict(list)
    if rider_col:
        for r in raw[rider_col].astype(str):
            if r.strip():
                variants[r.strip().lower()].append(r.strip())
    rows = [(next((v for v in vs if v != v.lower()), vs[0].title()), DEFAULT_CAPS.get(low, len(vs)))
            for low, vs in variants.items()]
    rows += [(d.title(), c) for d, c in DEFAULT_CAPS.items() if d not in variants]
    return pd.DataFrame({"Rider": [r for r, _ in rows], "Max Parcels": [c for _, c in rows],
                         "Available": True, "Backup Only": False, "Home Areas": ""})


def sample_template() -> bytes:
    df = pd.DataFrame({
        "Rider": ["Arif", "Arif", "", "", "Shah", "", "Fairuz", ""],
        "Order Number": ["A1001", "A1002", "A1003", "A1004", "A1005", "A1006", "A1007", "A1007"],
        "Delivery Address": [
            "12, Jalan SS2/3, 47300 Petaling Jaya, Selangor",
            "12 Jln SS2/3, 47300 Petaling Jaya",
            "8 Jalan Bukit Bintang, 55100 Kuala Lumpur",
            "No 5, Lorong Maarof, Bangsar, 59000 Kuala Lumpur",
            "22 Persiaran Kayangan, 40150 Shah Alam, Selangor",
            "A-12-3 Vista Apartment, Jalan Tun Razak, 50400 Kuala Lumpur",
            "Taman Tun Dr Ismail, 60000 Kuala Lumpur",
            "Taman Tun Dr Ismail, 60000 Kuala Lumpur"],
        "Remarks": ["Call 012-345 6789", "tel +60123456789", "", "Leave at guard 019 8887777",
                    "", "017-222 3333", "Call 0198887777", ""],
    })
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def check_admin() -> bool:
    pwd = os.environ.get("ADMIN_PASSWORD")
    if not pwd:
        try:
            pwd = st.secrets.get("ADMIN_PASSWORD")
        except Exception:
            pwd = None
    if not pwd or st.session_state.get("is_admin"):
        return True
    entered = st.text_input("Admin password", type="password")
    if entered:
        if hmac.compare_digest(entered, str(pwd)):
            st.session_state["is_admin"] = True
            st.rerun()
        st.error("Incorrect password.")
    return False


# --------------------------------------------------------------------------- #
# App
# --------------------------------------------------------------------------- #
def main():
    st.title("🚚 Delivery Order Router")
    st.caption("Upload orders → clean & flag duplicates → auto-assign riders with capacity caps.")
    if not check_admin():
        st.stop()

    with st.sidebar:
        st.header("1 · Upload")
        up = st.file_uploader("Delivery orders (.xlsx)", type=["xlsx", "xlsm"])

    if not up:
        st.info("Upload an Excel file with columns such as **Rider, Order Number, Delivery Address, Remarks**.")
        st.download_button("Download sample template", sample_template(), "sample_orders.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        return

    data = up.getvalue()
    sheets = pd.ExcelFile(io.BytesIO(data)).sheet_names
    with st.sidebar:
        sheet = st.selectbox("Sheet", sheets) if len(sheets) > 1 else sheets[0]
    with st.sidebar:
        header = st.checkbox("First row = column names", value=read_sheet(data, sheet, True).shape[1] > 1)
    raw = read_sheet(data, sheet, header)
    first_row = 2 if header else 1
    if raw.empty:
        st.error("The selected sheet is empty.")
        return
    cols = list(raw.columns)
    opt = ["(none)"] + cols

    with st.sidebar:
        st.header("2 · Columns")
        addr_c = st.selectbox("Delivery Address * (may hold name, phone, address, notes)", cols, index=cols.index(
            guess(cols, ["address", "addr"], cols[min(1, len(cols) - 1)])))
        order_c = st.selectbox("Order Number (optional)", opt, index=opt.index(
            guess(cols, ["order number", "order no", "order id", "order", "number"], "(none)")),
            help="Only values that look like order IDs are used; otherwise rows are referenced as 'Row N'.")
        remark_c = st.selectbox("Remarks (optional)", opt,
                                index=opt.index(guess(cols, ["remark", "note", "comment"], "(none)")))
        rider_c = st.selectbox("Rider", opt, index=opt.index(guess(cols, ["rider", "driver"], "(none)")))
        phone_c = st.selectbox("Phone (optional)", opt,
                               index=opt.index(guess(cols, ["phone", "tel", "mobile", "contact", "hp"], "(none)")))
        st.header("3 · Rules")
        threshold = st.slider("Fuzzy address threshold", 70, 100, 88,
                              help="Higher = stricter. 85-92 works well for typos and Jalan/Jln variants.")
        ignore_unit = st.checkbox("Ignore similar addresses with different unit/house numbers", True,
                                  help="Avoids flagging neighbours on the same street.")
        respect = st.radio("Existing Rider column", ["Keep pre-assigned (respect caps)", "Reassign everything"]) \
            .startswith("Keep")

    cols_sel = [None if c == "(none)" else c for c in (order_c, addr_c, remark_c, rider_c, phone_c)]
    if order_c == addr_c:
        st.error("Order Number and Delivery Address must be different columns.")
        return

    # ---- Rider roster & caps ------------------------------------------------
    sig = f"{up.name}|{len(data)}|{sheet}|{header}|{cols_sel[3]}"
    if st.session_state.get("roster_sig") != sig:
        st.session_state["roster_sig"] = sig
        st.session_state["roster_df"] = default_roster(raw, cols_sel[3])
        st.session_state.pop("result", None)

    st.subheader("Rider roster & capacity caps")
    st.caption("Caps start from your usual limits (Arif 15, Fairuz 14, Shah 18) or, for other riders, the number of "
               "orders they already have in the file - edit freely. Set each rider's max parcels. Untick *Available* for riders off today; tick *Backup Only* for "
               "riders who should only receive overflow. *Home Areas* (optional): comma-separated postcodes "
               "or area names, e.g. `47300, Bangsar`. Add rows for extra riders.")
    roster = st.data_editor(
        st.session_state["roster_df"], key=f"roster_{sig}", num_rows="dynamic", hide_index=True,
        use_container_width=True,
        column_config={
            "Max Parcels": st.column_config.NumberColumn(min_value=0, step=1, default=FALLBACK_CAP),
            "Available": st.column_config.CheckboxColumn(default=True),
            "Backup Only": st.column_config.CheckboxColumn(default=False),
            "Home Areas": st.column_config.TextColumn(),
        })

    total_cap = int(roster.loc[roster["Available"].fillna(True).astype(bool), "Max Parcels"].fillna(0).sum())
    st.caption(f"Orders in file: **{len(raw)}** · Total capacity of available riders: **{total_cap}**")

    if st.button("🚀 Clean, flag & auto-assign", type="primary"):
        if not roster["Rider"].dropna().astype(str).str.strip().any():
            st.error("Add at least one rider.")
        else:
            with st.spinner("Cleaning, matching and assigning..."):
                res = run_pipeline(raw, cols_sel, roster, threshold, ignore_unit, respect, first_row)
            st.session_state["result"] = res

    res = st.session_state.get("result")
    if not res:
        st.info("Configure the roster, then click **Clean, flag & auto-assign**.")
        return

    routing, flagged, loads, rc = res["routing"], res["flagged"], res["loads"], res["rider_col"]
    st.download_button("⬇️ Download updated routing sheet (.xlsx)", build_excel(res),
                       "routing_sheet.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    ok_ids, n_rows = res["order_ids"]
    if ok_ids < n_rows:
        st.info(f"Only {ok_ids} of {n_rows} rows have a usable order ID (the Number column mostly repeats rider "
                "names), so the rest are referenced by their sheet row number (e.g. 'Row 14').")

    tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "⚠️ Flagged duplicates", "🗺️ Routing sheet"])

    with tab1:
        unassigned = int((routing[rc] == "UNASSIGNED").sum())
        m = st.columns(5)
        m[0].metric("Orders", len(routing))
        m[1].metric("Flagged orders", int((routing["Flag Level"] != "").sum()))
        m[2].metric("Duplicate groups", routing.loc[routing["Dup Group"] != "", "Dup Group"].nunique())
        m[3].metric("Overflow reassigned", int((routing["Assignment Type"] == OVERFLOW).sum()))
        m[4].metric("Unassigned", unassigned)
        if unassigned:
            st.error(f"{unassigned} orders could not be assigned: every available rider is at cap. "
                     "Raise caps or add riders, then re-run.")
        st.subheader("Rider loads")
        st.dataframe(loads, hide_index=True, use_container_width=True, column_config={
            "Utilization %": st.column_config.ProgressColumn(min_value=0, max_value=100, format="%.0f%%")})
        st.bar_chart(loads.set_index("Rider")[["Assigned", "Free Slots"]])
        with st.expander("Orders per area by rider"):
            st.dataframe(pd.crosstab(routing["Area"].replace("", "(unknown)"), routing[rc]),
                         use_container_width=True)

    with tab2:
        if flagged.empty:
            st.success("No suspicious duplicates found.")
        else:
            f1, f2 = st.columns([1, 2])
            lv = f1.multiselect("Severity", ["High", "Medium", "Low"], ["High", "Medium", "Low"])
            q = f2.text_input("Search (order, address, phone, reason)")
            view = flagged[flagged["Flag Level"].isin(lv)]
            if q:
                mask = view.astype(str).apply(lambda s: s.str.contains(q, case=False, regex=False)).any(axis=1)
                view = view[mask]
            st.caption("High: duplicate order no. or same phone + similar address · Medium: similar address with "
                       "a different phone, or same phone at different addresses · Low: repeat phone / similar "
                       "address with no phone to compare.")
            st.dataframe(view, hide_index=True, use_container_width=True)

    with tab3:
        f1, f2 = st.columns([3, 1])
        rf = f1.multiselect("Filter by rider", sorted(routing[rc].unique()))
        full = f2.checkbox("Show all columns")
        view = routing[routing[rc].isin(rf)] if rf else routing
        if not full:
            view = view[["Source Row", "Order Ref", "Customer", "Parsed Address", "Postcode", "Area",
                         "Phone (Normalized)", rc, "Assignment Type", "Assignment Note",
                         "Flag Level", "Flag Reason", "Dup Group"]]
        st.dataframe(view, hide_index=True, use_container_width=True)


if __name__ == "__main__":
    main()
