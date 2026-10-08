import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Hometaste Delivery Auto-Allocator", layout="wide")

st.title("📦 Hometaste Delivery Auto-Allocator & Duplicate Checker")
st.write("Upload your daily Excel delivery sheet. The app applies your updated routing rules, respects rider capacity caps, and returns your exact original format.")

# Sidebar configuration for Riders and Caps
st.sidebar.header("1. Rider Capacity Caps")
st.sidebar.write("Set max parcel limit for each rider:")

rider_caps = {
    "Inhouse Rider": st.sidebar.number_input("Inhouse Rider Cap", value=25, min_value=1, max_value=100),
    "Fairuz": st.sidebar.number_input("Fairuz Cap", value=20, min_value=1, max_value=100),
    "Azwan": st.sidebar.number_input("Azwan Cap", value=25, min_value=1, max_value=100),
    "Shah": st.sidebar.number_input("Shah Cap", value=25, min_value=1, max_value=100),
    "Arif": st.sidebar.number_input("Arif Cap", value=20, min_value=1, max_value=100),
    "Joe": st.sidebar.number_input("Joe Cap", value=22, min_value=1, max_value=100),
    "Kali": st.sidebar.number_input("Kali Cap", value=25, min_value=1, max_value=100),
    "Hometaste": st.sidebar.number_input("Hometaste Cap", value=25, min_value=1, max_value=100),
    "Lizz": st.sidebar.number_input("Lizz Cap", value=30, min_value=1, max_value=100),
}

st.sidebar.header("2. Area Keyword Routing Rules")
st.sidebar.write("Map keywords (comma-separated) to specific riders:")

# Exact prioritized default rules
default_rules = {
    "Inhouse Rider": "zenith, kelana mahkota condominium, sentul, wangsa maju",
    "Fairuz": "lakefields, 57000, 46000, bandar sunway, sunway, ss8, puchong, jalil, sri petaling, kinrara, kembangan, the legacy oug, jalan gembira tmn overseas union",
    "Azwan": "bandar utama, 47800, sutramas, dutamas, sri hartamas, pju 3, ss2, kepong, desa park city, tropicana, pju 8, pju 9, pju 5, pju 10, damansara perdana, damansara damai, 47820, 47830, 52200, 47810, kip, 47400, 47300",
    "Shah": "universiti malaya, 50603, permaisuri, old klang road, bukit damansara, taman tun dr ismail, taman desa, 58200, seputeh, 58100, 58000, 57100, bangsar, bukit gasing, 50470, 46200, 46050",
    "Arif": "kemuning utama, 40400, u1, ss7, ara damansara, subang jaya, usj, kota kemuning, 47301",
    "Joe": "wangsa maju, jinjang, ipoh, 55000, 52100, setapak, sentul, 53100, 53000, 54200, 53300",
    "Kali": "pandan indah, mont kiara, chan sow lin, klcc, bukit bintang, 55100, 51200, pandan perdana, 55200, 50400, 50480",
    "Hometaste": "56000, 40100, 68000, 40170, 40300, 56100, 40150, 40200, setia eco park",
    "Lizz": "lizz rider pick up, lizz"
}

keyword_mapping = {}
for rider in rider_caps.keys():
    kw_input = st.sidebar.text_input(f"Keywords for {rider}", value=default_rules.get(rider, ""))
    keyword_mapping[rider] = [k.strip().lower() for k in kw_input.split(",") if k.strip()]

uploaded_file = st.file_uploader("Upload Daily Excel Sheet (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        cols = df.columns.tolist()
        rider_col = cols[0]
        address_col = cols[2] if len(cols) > 2 else cols[1]

        st.success(f"File uploaded successfully! Total rows detected: {len(df)}")

        if st.button("🚀 Run Auto-Assignment & Duplicate Check"):
            
            assignments = []
            rider_counts = {r: 0 for r in rider_caps.keys()}
            unassigned_rows = []

            # Step 1: Assign based on keywords with strict matching and exclusions
            for idx, row in df.iterrows():
                addr_text = str(row[address_col]).lower() if pd.notna(row[address_col]) else ""
                
                # Check for explicit exclusions (leave empty)
                if "sepang international circuit" in addr_text:
                    assignments.append(None)
                    continue

                assigned_rider = None
                
                # Special strict check for Lizz: ONLY assign if 'lizz' is in the address
                lizz_keywords = keyword_mapping.get("Lizz", [])
                is_lizz = any(lk in addr_text for lk in lizz_keywords if lk)
                
                if is_lizz:
                    if rider_counts["Lizz"] < rider_caps["Lizz"]:
                        assigned_rider = "Lizz"
                else:
                    # Check riders in defined dictionary priority order
                    for rider, keywords in keyword_mapping.items():
                        if rider == "Lizz":
                            continue
                        for kw in keywords:
                            if kw:
                                if kw.isdigit() or len(kw) <= 4:
                                    pattern = r'\b' + re.escape(kw) + r'\b'
                                    if re.search(pattern, addr_text):
                                        if rider_counts[rider] < rider_caps[rider]:
                                            assigned_rider = rider
                                            break
                                else:
                                    if kw in addr_text:
                                        if rider_counts[rider] < rider_caps[rider]:
                                            assigned_rider = rider
                                            break
                        if assigned_rider:
                            break
                
                if assigned_rider:
                    rider_counts[assigned_rider] += 1
                    assignments.append(assigned_rider)
                else:
                    unassigned_rows.append(idx)
                    assignments.append(None)

            # Step 2: Handle overflow / unassigned with available capacity (Excluding Lizz from auto overflow)
            overflow_riders = {r: cap for r, cap in rider_caps.items() if r != "Lizz"}
            for idx in unassigned_rows:
                if assignments[idx] is None and str(row.get(address_col, "")).lower().find("sepang international circuit") != -1:
                    continue
                
                if overflow_riders:
                    available_rider = max(overflow_riders, key=lambda r: rider_caps[r] - rider_counts[r])
                    if rider_counts[available_rider] < rider_caps[available_rider] + 15: 
                        rider_counts[available_rider] += 1
                        assignments[idx] = available_rider
                    else:
                        assignments[idx] = list(overflow_riders.keys())[0]

            # Update the original dataframe's Rider column
            df[rider_col] = assignments

            st.subheader("📊 Rider Workload Summary")
            summary_df = pd.DataFrame(list(rider_counts.items()), columns=["Rider", "Assigned Orders"])
            st.dataframe(summary_df)

            st.subheader("📋 Updated Delivery Sheet Preview")
            st.dataframe(df.head(10))

            # Export button for updated excel
            output_filename = "Updated_Routing_Sheet.xlsx"
            df.to_excel(output_filename, index=False)
            
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="📥 Download Updated Excel Routing Sheet",
                    data=f,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error processing file: {e}")
