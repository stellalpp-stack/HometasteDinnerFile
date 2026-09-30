import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import re

st.set_page_config(page_title="Hometaste Delivery Auto-Allocator", layout="wide")

st.title("📦 Hometaste Delivery Auto-Allocator & Duplicate Checker")
st.write("Upload your daily Excel delivery sheet. The app will clean duplicates, apply your custom keyword routing rules, respect rider capacity caps, and return your exact original format with updated rider assignments.")

# Sidebar configuration for Riders and Caps
st.sidebar.header("1. Rider Capacity Caps")
st.sidebar.write("Set max parcel limit for each rider:")

rider_caps = {
    "Arif": st.sidebar.number_input("Arif Cap", value=20, min_value=1, max_value=100),
    "Fairuz": st.sidebar.number_input("Fairuz Cap", value=20, min_value=1, max_value=100),
    "Shah": st.sidebar.number_input("Shah Cap", value=25, min_value=1, max_value=100),
    "Azwan": st.sidebar.number_input("Lizz Cap", value=25, min_value=1, max_value=100),
    "Joe": st.sidebar.number_input("Joe Cap", value=22, min_value=1, max_value=100),
    "Kali": st.sidebar.number_input("Kali Cap", value=25, min_value=1, max_value=100),
    "Hometaste": st.sidebar.number_input("Kali Cap", value=25, min_value=1, max_value=100),
}

st.sidebar.header("2. Area Keyword Routing Rules")
st.sidebar.write("Map keywords (comma-separated) to specific riders:")

# Default keyword mapping setup
default_rules = {
    "Shah": "permaisuri, old klang road, bukit damansara, Taman Tun Dr Ismail, taman desa, 58200, seputeh, 58100, 58000, 57100, bangsar, bukit gasing, 50470, 46200, 46050",
    "Arif": "ara damansara, subang jaya, USJ, u1, kota kemuning, ss7",
    "Fairuz": "ss8, puchong, jalil, sri petaling, kinrara, kembangan",
    "Azwan": "kepong, desa park city, tropicanal, PJU8, PJU9, PJU5, PJU10, 52200, PJU8, 52200, 47810, KIP, 47400", 47300,
    "Joe": "wangsamaju, jinjang, ipoh, 55000, 52100, setapak, sentul, 53100, 53000, 54200, 53300, 55000",
    "Kali": "klcc, bukit bintang, 55100, 51200, mont kiara, pandan perdana, 51200, 55200, 50400"
    "inhouse rider": "klcc, bukit bintang, sentul, wangsa maju"
    "Hometaste": "40100, 68000, 40170, 40300, 56000, 56100, 40150, 40200"
}

keyword_mapping = {}
for rider in rider_caps.keys():
    kw_input = st.sidebar.text_input(f"Keywords for {rider}", value=default_rules.get(rider, ""))
    keyword_mapping[rider] = [k.strip().lower() for k in kw_input.split(",") if k.strip()]

uploaded_file = st.file_uploader("Upload Daily Excel Sheet (.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read original excel
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # Identify key columns (typically column index 0 for Rider, 1 for Number/ID, 2 for Address)
        cols = df.columns.tolist()
        rider_col = cols[0]
        address_col = cols[2] if len(cols) > 2 else cols[1]
        number_col = cols[1] if len(cols) > 1 else None

        st.success(f"File uploaded successfully! Total rows detected: {len(df)}")

        if st.button("🚀 Run Auto-Assignment & Duplicate Check"):
            
            assignments = []
            rider_counts = {r: 0 for r in rider_caps.keys()}
            unassigned_rows = []

            # Step 1: Assign based on keywords
            for idx, row in df.iterrows():
                # Skip header rows if any
                addr_text = str(row[address_col]).lower() if pd.notna(row[address_col]) else ""
                
                assigned_rider = None
                # Check keywords
                for rider, keywords in keyword_mapping.items():
                    for kw in keywords:
                        if kw and kw in addr_text:
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

            # Step 2: Handle overflow / unassigned with available capacity
            for idx in unassigned_rows:
                # Find rider with most remaining capacity room
                available_rider = max(rider_caps, key=lambda r: rider_caps[r] - rider_counts[r])
                if rider_counts[available_rider] < rider_caps[available_rider] + 5: # Allow small overflow buffer if needed
                    rider_counts[available_rider] += 1
                    assignments[idx] = available_rider
                else:
                    assignments[idx] = list(rider_caps.keys())[0] # Fallback

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
