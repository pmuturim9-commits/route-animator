import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Route Automator", layout="centered")

st.title("🚌 Route Summary Automator")

# Master Route Code Mapping
ROUTE_MAPPING = {
    "GITHARI": ("T65", 2), "BRIDGES": ("T04", 3), "CENTRE": ("T33", 4),
    "CHEBA NGURUKA": ("T05", 5), "CHOBE": ("T14", 7), "CHUMA": ("T09", 8),
    "DAM": ("T18", 10), "EXLEES": ("T68", 11), "FARU": ("T16", 12),
    "GATITU": ("T43", 16), "GITHIMA": ("T29", 17), "GITIRI": ("T31", 18),
    "KIBORE": ("T41", 38), "KIRIMA": ("T61", 43), "MAIN": ("T87", 46),
    "SEMIHEADQUATER": ("T59", 63)
}

mode = st.radio("Choose Mode:", ["Stage 1: Generate Weekly Report", "Stage 2: Generate Monthly Report from 4 Weeks"])

# ---------------------------------------------------------
# STAGE 1: WEEKLY REPORT
# ---------------------------------------------------------
if mode == "Stage 1: Generate Weekly Report":
    st.subheader("📅 Stage 1: Generate Single Weekly Summary")
    
    week_label = st.selectbox("Select Week:", ["Week 1", "Week 2", "Week 3", "Week 4"])
    uploaded_daily = st.file_uploader("Upload Raw Daily Excel File for this week", type=["xlsx", "xls"], key="daily")

    if uploaded_daily is not None:
        if st.button("Process Weekly Report"):
            try:
                df_raw = pd.read_excel(uploaded_daily)
                
                route_col = df_raw.columns[0]
                day_cols = [c for c in df_raw.columns[1:] if "Unnamed" not in str(c)]
                
                # Keep ALL unique routes including nil/empty ones
                all_routes = df_raw[route_col].dropna().unique()
                
                weekly_data = []
                for route in all_routes:
                    sub_df = df_raw[df_raw[route_col] == route]
                    # Convert values to numeric and replace blank/nil entries with 0
                    val_sum = sub_df[day_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum().sum()
                    weekly_data.append({"Route": str(route).strip(), week_label: val_sum})
                
                df_weekly = pd.DataFrame(weekly_data)
                
                st.success(f"{week_label} report created! All routes included.")
                st.dataframe(df_weekly, use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_weekly.to_excel(writer, sheet_name=week_label, index=False)
                
                st.download_button(
                    label=f"📥 Download {week_label} Summary (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"{week_label}_Summary.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error processing file: {e}")

# ---------------------------------------------------------
# STAGE 2: MONTHLY CONSOLIDATED REPORT
# ---------------------------------------------------------
else:
    st.subheader("📊 Stage 2: Combine 4 Weekly Reports into Final Monthly Summary")
    
    uploaded_weeks = st.file_uploader(
        "Upload all 4 Weekly Summary Files (.xlsx)", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True,
        key="monthly"
    )
    rate_multiplier = st.number_input("Enter Rate Multiplier for this month:", value=60.0, step=1.0)

    if uploaded_weeks:
        if len(uploaded_weeks) != 4:
            st.warning(f"You have uploaded {len(uploaded_weeks)} file(s). Please upload exactly 4 files (Week 1 to 4).")
        else:
            if st.button("Generate Final Monthly Report"):
                try:
                    combined_df = None
                    
                    # Outer join ensures all routes across all 4 weeks are present
                    for file in sorted(uploaded_weeks, key=lambda x: x.name):
                        df_wk = pd.read_excel(file)
                        if combined_df is None:
                            combined_df = df_wk
                        else:
                            combined_df = pd.merge(combined_df, df_wk, on="Route", how="outer")
                    
                    # Replace any NaN/missing weekly figures with 0 (nil)
                    week_cols = [c for c in combined_df.columns if c != "Route"]
                    combined_df[week_cols] = combined_df[week_cols].fillna(0)
                    
                    # Calculations
                    combined_df["Total Sum"] = combined_df[week_cols].sum(axis=1)
                    combined_df["AMOUNT"] = combined_df["Total Sum"] * rate_multiplier
                    
                    # Map Codes & Page Numbers
                    codes, pgs = [], []
                    for r in combined_df["Route"]:
                        info = ROUTE_MAPPING.get(str(r).strip().upper(), ("T00", 1))
                        codes.append(info[0])
                        pgs.append(info[1])
                        
                    combined_df["Code"] = codes
                    combined_df["PG. NO"] = pgs
                    
                    # Bottom Total Row across all columns
                    total_row = {
                        "Route": "SUM TOTAL",
                        "Total Sum": combined_df["Total Sum"].sum(),
                        "AMOUNT": combined_df["AMOUNT"].sum(),
                        "Code": "",
                        "PG. NO": ""
                    }
                    for wc in week_cols:
                        total_row[wc] = combined_df[wc].sum()
                        
                    df_final = pd.concat([combined_df, pd.DataFrame([total_row])], ignore_index=True)
                    
                    st.success("Final Monthly Report compiled! Includes all routes and nil figures.")
                    st.dataframe(df_final, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, sheet_name="Final Summary", index=False)
                        
                    st.download_button(
                        label="📥 Download Final Monthly Report (.xlsx)",
                        data=output.getvalue(),
                        file_name="Final_Monthly_Summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Error compiling monthly report: {e}")
