import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Route Summary Master", layout="wide")

st.title("🚌 Monthly Route Summary Automator")

# Master Route Codes
ROUTE_MAPPING = {
    "GITHARI": "T65", "BRIDGES": "T04", "CENTRE": "T33",
    "CHEBA NGURUKA": "T05", "CHEBA-B": "T57", "CHOBE": "T14",
    "CHUMA": "T09", "CIONDO-B": "T13", "DAM": "T18",
    "EXLEES": "T68", "FARU": "T16", "GACATA": "T40",
    "GACHUCHA-B": "T46", "GATHARA": "T67", "GATITU": "T43",
    "GITHIMA": "T29", "GITIRI": "T31", "GITITE": "T12",
    "GURD": "T52", "KIBORE": "T41", "KIRIMA": "T61",
    "MAIN": "T87", "SEMIHEADQUATER": "T59", "THINDI": "T20",
    "UPPER CIONDO": "T38", "WANGU": "T82", "YAANGA": "T27",
    "KWARE": "T90", "KARIMA": "T91", "MUTAMAIYU": "T39",
    "MUTAMAIYU PM": "T72", "CHUMA-B": "T95"
}

st.markdown("""
**Instructions:** Upload your current Master Monthly Workbook (optional for Week 1) and the raw Daily Excel file for the week you want to add/update.
""")

col1, col2 = st.columns([1, 1])

with col1:
    master_file = st.file_uploader("📁 Step 1: Upload Existing Master File (Leave empty for Week 1)", type=["xlsx", "xls"], key="master")

with col2:
    selected_week = st.selectbox("📅 Step 2: Select Week to Add/Update:", ["week 1", "week 2", "week 3", "week 4"])
    daily_file = st.file_uploader(f"📄 Step 3: Upload Raw Daily Excel File for {selected_week}", type=["xlsx", "xls"], key="daily")
    rate_multiplier = st.number_input("Rate Multiplier per Unit:", value=60.0, step=1.0)

if daily_file is not None:
    if st.button("🚀 Process & Update Master Workbook", type="primary"):
        try:
            # 1. Parse Daily File to Extract Weekly Totals & Daily Grid
            xls_daily = pd.ExcelFile(daily_file)
            df_raw = pd.read_excel(xls_daily, sheet_name=0)
            
            days = []
            day_col_indices = []
            for c_idx in range(len(df_raw.columns)):
                val = df_raw.iloc[0, c_idx]
                if pd.notna(val) and str(val).strip().isdigit():
                    days.append(f"Day {int(val)}")
                    day_col_indices.append(c_idx)
            
            route_records = {}
            for day_label, c_idx in zip(days, day_col_indices):
                route_col = c_idx - 1
                var_col = c_idx
                for r in range(2, len(df_raw)):
                    r_name = df_raw.iloc[r, route_col]
                    r_var = df_raw.iloc[r, var_col]
                    if pd.notna(r_name) and str(r_name).strip() != "" and str(r_name).strip() != "0":
                        r_clean = str(r_name).strip().upper()
                        if r_clean not in route_records:
                            route_records[r_clean] = {}
                        v_num = pd.to_numeric(r_var, errors='coerce') if pd.notna(r_var) else 0
                        route_records[r_clean][day_label] = float(v_num) if pd.notna(v_num) else 0.0

            df_daily_grid = pd.DataFrame.from_dict(route_records, orient='index').fillna(0)
            df_daily_grid.index.name = "Route"
            df_daily_grid = df_daily_grid.reset_index()
            day_cols_sorted = sorted([c for c in df_daily_grid.columns if c != "Route"], key=lambda x: int(x.replace("Day ", "")))
            df_daily_grid = df_daily_grid[["Route"] + day_cols_sorted].sort_values("Route")
            
            weekly_totals = df_daily_grid.set_index("Route")[day_cols_sorted].sum(axis=1).to_dict()

            # 2. Load or Initialize Master Table
            existing_sheets = {}
            if master_file is not None:
                xls_master = pd.ExcelFile(master_file)
                for s in xls_master.sheet_names:
                    existing_sheets[s] = pd.read_excel(xls_master, sheet_name=s)
                
                if "Monthly Summary" in existing_sheets:
                    df_master = existing_sheets["Monthly Summary"]
                else:
                    df_master = existing_sheets[xls_master.sheet_names[0]]
            else:
                df_master = pd.DataFrame(columns=["Route", "week 1", "week 2", "week 3", "week 4", "sum", "AMOUNT", "Code"])

            # Clean header if needed
            if len(df_master) > 0 and df_master.iloc[0]["Route"] == "Route":
                df_master.columns = df_master.iloc[0].values
                df_master = df_master.iloc[1:].reset_index(drop=True)

            # Strip out previous SUM TOTAL row
            df_master = df_master[df_master["Route"].astype(str).str.strip().str.upper() != "SUM TOTAL"].copy()

            # Ensure all standard week columns exist
            for col in ["week 1", "week 2", "week 3", "week 4"]:
                if col not in df_master.columns:
                    df_master[col] = 0.0
                df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0.0)

            # Merge / Update route figures
            route_series = df_master["Route"].astype(str).str.strip().str.upper()
            all_routes = set(route_series).union(set(weekly_totals.keys()))
            master_dict = df_master.set_index(df_master["Route"].astype(str).str.strip().str.upper()).to_dict('index')
            
            updated_rows = []
            for r in sorted(all_routes):
                row_data = master_dict.get(r, {"Route": r, "week 1": 0.0, "week 2": 0.0, "week 3": 0.0, "week 4": 0.0})
                
                # Insert / Overwrite figure for selected week
                if r in weekly_totals:
                    row_data[selected_week] = weekly_totals[r]
                
                w_sum = sum([float(row_data.get(w, 0.0)) for w in ["week 1", "week 2", "week 3", "week 4"]])
                amt = w_sum * rate_multiplier
                code = ROUTE_MAPPING.get(r, "T00")
                
                row_data["Route"] = row_data.get("Route", r)
                row_data["sum"] = w_sum
                row_data["AMOUNT"] = amt
                row_data["Code"] = code
                updated_rows.append(row_data)

            df_updated_master = pd.DataFrame(updated_rows)
            cols_order = ["Route", "week 1", "week 2", "week 3", "week 4", "sum", "AMOUNT", "Code"]
            df_updated_master = df_updated_master[cols_order]

            # Append SUM TOTAL row
            total_row = {
                "Route": "SUM TOTAL",
                "week 1": df_updated_master["week 1"].sum(),
                "week 2": df_updated_master["week 2"].sum(),
                "week 3": df_updated_master["week 3"].sum(),
                "week 4": df_updated_master["week 4"].sum(),
                "sum": df_updated_master["sum"].sum(),
                "AMOUNT": df_updated_master["AMOUNT"].sum(),
                "Code": ""
            }
            df_final_master = pd.concat([df_updated_master, pd.DataFrame([total_row])], ignore_index=True)

            # Keep sheets updated
            existing_sheets["Monthly Summary"] = df_final_master
            existing_sheets[f"{selected_week.capitalize()} Daily"] = df_daily_grid

            st.success(f"Successfully updated **{selected_week}** in the Master Workbook!")

            tab1, tab2 = st.tabs(["📊 Cumulative Monthly Summary", f"📅 {selected_week.capitalize()} Daily Breakdown"])
            with tab1:
                st.dataframe(df_final_master, use_container_width=True)
            with tab2:
                st.dataframe(df_daily_grid, use_container_width=True)

            # Export Excel workbook with all cumulative sheets
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final_master.to_excel(writer, sheet_name="Monthly Summary", index=False)
                for s_name, s_df in existing_sheets.items():
                    if s_name != "Monthly Summary":
                        s_df.to_excel(writer, sheet_name=s_name, index=False)

            st.download_button(
                label="📥 Download Updated Cumulative Master Workbook (.xlsx)",
                data=output.getvalue(),
                file_name="Monthly_Route_Summary_Master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred while processing: {e}")
