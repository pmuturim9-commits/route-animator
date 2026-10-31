import io
import re
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Route Summary Master Pro", layout="wide")
st.title("🚌 Monthly Route Summary Automator Pro")


def clean_str(val):
    """Normalize string inputs by stripping extra spaces and converting to uppercase."""
    if pd.isna(val):
        return ""
    cleaned = re.sub(r"\s+", " ", str(val)).strip().upper()
    return cleaned


def parse_custom_codes_file(uploaded_file):
    """Dynamically parses route-to-code mapping sheets."""
    df_raw = pd.read_excel(uploaded_file)

    start_row = 0
    code_col_idx = 0
    route_col_idx = 1
    found_header = False

    for idx, row in df_raw.iterrows():
        row_str = [clean_str(val) for val in row.values if pd.notna(val)]
        if any("CODE" in s for s in row_str) or any("ROUTE" in s for s in row_str):
            start_row = idx + 1
            found_header = True
            for col_i, val in enumerate(row.values):
                val_s = clean_str(val)
                if "CODE" in val_s:
                    code_col_idx = col_i
                elif "ROUTE" in val_s:
                    route_col_idx = col_i
            break

    if not found_header:
        code_col_idx = 0
        route_col_idx = 1 if len(df_raw.columns) > 1 else 0
        start_row = 0

    custom_mapping = {}
    for idx in range(start_row, len(df_raw)):
        if idx >= len(df_raw):
            break
        code_val = df_raw.iloc[idx, code_col_idx] if code_col_idx < len(df_raw.columns) else ""
        route_val = df_raw.iloc[idx, route_col_idx] if route_col_idx < len(df_raw.columns) else ""

        if pd.notna(route_val) and pd.notna(code_val):
            r_str = clean_str(route_val)
            c_str = clean_str(code_val)

            if r_str not in [
                "ROUTE",
                "ROUTE NAME",
                "NAN",
                "",
                "NONE",
            ] and c_str not in ["CODE", "T-CODE", "TCODE", "NAN", ""]:
                custom_mapping[r_str] = c_str

    return custom_mapping


def style_excel_workbook(writer, df_master):
    """Applies clean formatting and cell styling to output workbook safely."""
    header_fill = PatternFill(
        start_color="0D9488", end_color="0D9488", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    total_fill = PatternFill(
        start_color="FEF08A", end_color="FEF08A", fill_type="solid"
    )
    bold_font = Font(name="Calibri", size=11, bold=True)

    # 1. Style Monthly Summary
    if "Monthly Summary" in writer.sheets:
        ws_master = writer.sheets["Monthly Summary"]
        for col_idx in range(1, len(df_master.columns) + 1):
            cell = ws_master.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row_idx in range(2, len(df_master) + 2):
            is_total_row = row_idx == len(df_master) + 1
            for col_idx in range(1, len(df_master.columns) + 1):
                cell = ws_master.cell(row=row_idx, column=col_idx)
                if is_total_row:
                    cell.fill = total_fill
                    cell.font = bold_font

                col_name = df_master.columns[col_idx - 1]
                if col_name in [
                    "week 1",
                    "week 2",
                    "week 3",
                    "week 4",
                    "Total Shortages Vol",
                    "Collected Volume",
                ]:
                    cell.number_format = "#,##0.0"
                elif col_name in [
                    "Shortage Amount",
                    "Gross Earning",
                    "Other Deductions",
                    "Total Book Dr",
                    "Net Earning",
                ]:
                    cell.number_format = "KSh #,##0.00"

    # 2. Style Net Earnings Analysis Sheet
    if "Net Earnings Analysis" in writer.sheets:
        ws_earnings = writer.sheets["Net Earnings Analysis"]
        for row in range(1, ws_earnings.max_row + 1):
            is_last_row = row == ws_earnings.max_row
            for col in range(1, ws_earnings.max_column + 1):
                cell = ws_earnings.cell(row=row, column=col)
                if row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif is_last_row:
                    cell.fill = total_fill
                    cell.font = bold_font

                if col in [3, 7]:  # Net Earning columns
                    cell.number_format = "KSh #,##0.00"

    # 3. Style Daily sheets & Safely auto-adjust column widths
    for sheetname in writer.sheets:
        ws_curr = writer.sheets[sheetname]
        if "Daily" in sheetname:
            max_r = ws_curr.max_row
            max_c = ws_curr.max_column
            for c in range(1, max_c + 1):
                ws_curr.cell(row=max_r, column=c).fill = total_fill
                ws_curr.cell(row=max_r, column=c).font = bold_font
            for r in range(1, max_r + 1):
                ws_curr.cell(row=r, column=max_c).font = bold_font

        for col in ws_curr.columns:
            max_len = 0
            for cell in col:
                val_str = str(cell.value or "")
                line_lens = [len(line) for line in val_str.split("\n")]
                if line_lens:
                    max_len = max(max_len, max(line_lens))
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws_curr.column_dimensions[col_letter].width = max(max_len + 4, 14)


# --- UI CONTROLS ---
col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    codes_file = st.file_uploader(
        "🏷️ Step 1: Upload Route Codes Sheet (Required)",
        type=["xlsx", "xls"],
        key="codes",
    )
    master_file = st.file_uploader(
        "📁 Step 2: Upload Existing Master File (Optional)",
        type=["xlsx", "xls"],
        key="master",
    )

with col2:
    selected_week = st.selectbox(
        "📅 Step 3: Select Week to Update:",
        ["week 1", "week 2", "week 3", "week 4"],
    )
    daily_file = st.file_uploader(
        f"📄 Step 4: Upload Raw Daily File ({selected_week})",
        type=["xlsx", "xls"],
        key="daily",
    )
    shortage_rate = st.number_input(
        "Shortage Rate per Unit (KSh):", value=60.0, step=1.0
    )

with col3:
    st.markdown("### 💰 Step 5: Earnings & Extra Deductions")
    monthly_collected_file = st.file_uploader(
        "📄 Upload Monthly Collected Totals File (Optional)",
        type=["xlsx", "xls"],
        key="collected",
    )
    earning_rate = st.number_input(
        "Gross Earning Rate Multiplier (KSh):", value=70.0, step=1.0
    )

active_route_mapping = {}

if codes_file is not None:
    try:
        active_route_mapping = parse_custom_codes_file(codes_file)
        st.success(f"✅ Loaded {len(active_route_mapping)} route codes!")
    except Exception as e:
        st.error(f"⚠️ Error reading Route Codes file: {e}")
else:
    st.info("📌 Please upload your Route Codes file in Step 1 to begin.")

# Deductions Management Section
st.markdown("---")
st.subheader("➕ Custom Route Deductions Management")

if "other_deductions" not in st.session_state:
    st.session_state.other_deductions = {}

ded_tab1, ded_tab2 = st.tabs(["Manual Deduction Input", "Bulk Deductions File Upload"])

with ded_tab1:
    if active_route_mapping:
        route_list = sorted(list(active_route_mapping.keys()))
        ded_col1, ded_col2, ded_col3 = st.columns([2, 2, 1])

        with ded_col1:
            route_to_deduct = st.selectbox("Select Route for Custom Deduction:", route_list)
        with ded_col2:
            ded_amount = st.number_input("Deduction Amount (KSh):", min_value=0.0, step=100.0)
        with ded_col3:
            st.write(" ")
            st.write(" ")
            if st.button("Add Deduction"):
                if route_to_deduct:
                    st.session_state.other_deductions[route_to_deduct] = round(
                        st.session_state.other_deductions.get(route_to_deduct, 0.0) + ded_amount, 2
                    )
                    st.success(f"Added KSh {ded_amount:,.2f} deduction to {route_to_deduct}")

with ded_tab2:
    ded_file = st.file_uploader("Upload Bulk Deductions (CSV/Excel with 'Route' and 'Deduction' columns):", type=["csv", "xlsx"])
    if ded_file is not None:
        try:
            df_ded_in = pd.read_csv(ded_file) if ded_file.name.endswith(".csv") else pd.read_excel(ded_file)
            
            # Dynamically locate header row containing 'ROUTE'
            start_row = 0
            r_col_idx = None
            v_col_idx = None
            
            for idx, row in df_ded_in.iterrows():
                row_str = [clean_str(val) for val in row.values if pd.notna(val)]
                if any("ROUTE" in s for s in row_str):
                    start_row = idx + 1
                    for col_i, val in enumerate(row.values):
                        val_s = clean_str(val)
                        if "ROUTE" in val_s:
                            r_col_idx = col_i
                        elif "DEDUCT" in val_s or "AMOUNT" in val_s:
                            v_col_idx = col_i
                    break
            
            # Fallback if headers are in standard columns
            if r_col_idx is None:
                for c_idx, col in enumerate(df_ded_in.columns):
                    if "ROUTE" in clean_str(col):
                        r_col_idx = c_idx
                    elif "DEDUCT" in clean_str(col) or "AMOUNT" in clean_str(col):
                        v_col_idx = c_idx

            r_col_idx = r_col_idx if r_col_idx is not None else 1
            v_col_idx = v_col_idx if v_col_idx is not None else 2

            r_col = df_ded_in.columns[r_col_idx]
            v_col = df_ded_in.columns[v_col_idx]

            df_ded_clean = df_ded_in.iloc[start_row:] if start_row > 0 else df_ded_in

            loaded_count = 0
            for _, r_item in df_ded_clean.iterrows():
                r_c = clean_str(r_item[r_col])
                v_c = pd.to_numeric(r_item[v_col], errors="coerce")
                if r_c and pd.notna(v_c) and r_c not in ["ROUTE", "NAN", "", "NONE"]:
                    st.session_state.other_deductions[r_c] = round(float(v_c), 2)
                    loaded_count += 1
                    
            st.success(f"Successfully loaded {loaded_count} bulk deductions!")
        except Exception as e:
            st.error(f"Error parsing deductions file: {e}")

if st.session_state.other_deductions:
    with st.expander("View Active Custom Deductions"):
        ded_df = pd.DataFrame(
            list(st.session_state.other_deductions.items()),
            columns=["Route", "Other Deductions (KSh)"],
        )
        st.dataframe(ded_df, use_container_width=True)
        if st.button("Clear All Deductions"):
            st.session_state.other_deductions = {}
            st.rerun()

st.markdown("---")

if daily_file is not None:
    if not active_route_mapping:
        st.error("❌ You must upload a Route Codes file in Step 1 before processing!")
    elif st.button("🚀 Process & Update Master Workbook", type="primary"):
        try:
            # 1. Parse Daily File
            xls_daily = pd.ExcelFile(daily_file)
            df_raw = pd.read_excel(xls_daily, sheet_name=0)

            day_cols_info = []
            header_has_days = any(str(c).strip().isdigit() for c in df_raw.columns)

            if header_has_days:
                for c_idx, col_name in enumerate(df_raw.columns):
                    c_str = str(col_name).strip()
                    if c_str.isdigit() and c_idx > 0:
                        day_cols_info.append((f"Day {int(c_str)}", c_idx - 1, c_idx))
                start_row = 0
            else:
                for c_idx in range(len(df_raw.columns)):
                    val = df_raw.iloc[0, c_idx]
                    if pd.notna(val) and str(val).strip().isdigit() and c_idx > 0:
                        day_cols_info.append(
                            (f"Day {int(str(val).strip())}", c_idx - 1, c_idx)
                        )
                start_row = 1

            route_records = {}
            unmapped_routes = set()
            for day_label, route_col, val_col in day_cols_info:
                for r in range(start_row, len(df_raw)):
                    r_name = df_raw.iloc[r, route_col] if route_col < len(df_raw.columns) else None
                    r_var = df_raw.iloc[r, val_col] if val_col < len(df_raw.columns) else None
                    if pd.notna(r_name) and clean_str(r_name) not in ["", "0", "ROUTE", "NAN"]:
                        r_clean = clean_str(r_name)
                        if r_clean not in active_route_mapping:
                            unmapped_routes.add(r_clean)
                        if r_clean not in route_records:
                            route_records[r_clean] = {}
                        v_num = pd.to_numeric(r_var, errors="coerce") if pd.notna(r_var) else 0.0
                        route_records[r_clean][day_label] = float(v_num) if pd.notna(v_num) else 0.0

            if unmapped_routes:
                st.warning(f"⚠️ Unmapped Routes Detected: {', '.join(unmapped_routes)}. Default code 'T00' assigned.")

            # 2. Build Daily Grid
            df_daily_grid = pd.DataFrame.from_dict(route_records, orient="index").fillna(0.0)
            df_daily_grid.index.name = "Route"
            df_daily_grid = df_daily_grid.reset_index()

            day_cols_sorted = sorted(
                [c for c in df_daily_grid.columns if c != "Route"],
                key=lambda x: int(x.replace("Day ", "")),
            )
            df_daily_grid = df_daily_grid[["Route"] + day_cols_sorted].sort_values("Route")
            weekly_totals = (
                df_daily_grid.set_index("Route")[day_cols_sorted].sum(axis=1).round(2).to_dict()
            )

            df_daily_grid["Code"] = df_daily_grid["Route"].apply(
                lambda x: active_route_mapping.get(clean_str(x), "T00")
            )
            df_daily_grid["Total"] = df_daily_grid[day_cols_sorted].sum(axis=1).round(2)

            final_daily_cols = ["Route", "Code"] + day_cols_sorted + ["Total"]
            df_daily_grid = df_daily_grid[final_daily_cols]

            daily_totals_row = {"Route": "SUM TOTAL", "Code": ""}
            for col in day_cols_sorted + ["Total"]:
                daily_totals_row[col] = round(float(df_daily_grid[col].sum()), 2)

            df_daily_grid_final = pd.concat(
                [df_daily_grid, pd.DataFrame([daily_totals_row])],
                ignore_index=True,
            )

            # 3. Parse Monthly Collected Totals
            collected_dict = {}
            if monthly_collected_file is not None:
                df_coll_raw = pd.read_excel(monthly_collected_file)
                r_col = df_coll_raw.columns[0]
                v_col = df_coll_raw.columns[1] if len(df_coll_raw.columns) > 1 else df_coll_raw.columns[0]

                for _, row in df_coll_raw.iterrows():
                    r_val = clean_str(row[r_col])
                    if pd.notna(row[r_col]) and r_val not in ["ROUTE", "NAN", ""]:
                        val_num = pd.to_numeric(row[v_col], errors="coerce")
                        collected_dict[r_val] = round(float(val_num), 2) if pd.notna(val_num) else 0.0

            # 4. Process Master Summary safely
            existing_sheets = {}
            if master_file is not None:
                xls_master = pd.ExcelFile(master_file)
                for s in xls_master.sheet_names:
                    if s not in ["Monthly Summary", "Net Earnings Analysis"]:
                        existing_sheets[s] = pd.read_excel(xls_master, sheet_name=s)
                df_master = pd.read_excel(xls_master, sheet_name="Monthly Summary") if "Monthly Summary" in xls_master.sheet_names else pd.read_excel(xls_master, sheet_name=0)
            else:
                df_master = pd.DataFrame(
                    columns=["Route", "week 1", "week 2", "week 3", "week 4"]
                )

            if "Route" in df_master.columns and len(df_master) > 0:
                if str(df_master.iloc[0]["Route"]).strip().upper() == "ROUTE":
                    df_master.columns = df_master.iloc[0].values
                    df_master = df_master.iloc[1:].reset_index(drop=True)

            if "Route" in df_master.columns:
                df_master = df_master[
                    df_master["Route"].astype(str).str.strip().str.upper() != "SUM TOTAL"
                ].copy()
                df_master["Route_Clean"] = df_master["Route"].apply(clean_str)
            else:
                df_master["Route_Clean"] = []

            for col in ["week 1", "week 2", "week 3", "week 4"]:
                if col not in df_master.columns:
                    df_master[col] = 0.0
                df_master[col] = pd.to_numeric(df_master[col], errors="coerce").fillna(0.0)

            if not df_master.empty:
                df_master = df_master.groupby("Route_Clean", as_index=False).agg({
                    "Route": "first",
                    "week 1": "sum",
                    "week 2": "sum",
                    "week 3": "sum",
                    "week 4": "sum",
                })

            all_routes = sorted(
                list(
                    set(df_master["Route_Clean"] if "Route_Clean" in df_master else [])
                    .union(set(weekly_totals.keys()))
                    .union(set(collected_dict.keys()))
                )
            )
            master_dict = df_master.set_index("Route_Clean").to_dict("index") if not df_master.empty else {}

            updated_rows = []
            for r in all_routes:
                row_data = master_dict.get(
                    r,
                    {"Route": r, "week 1": 0.0, "week 2": 0.0, "week 3": 0.0, "week 4": 0.0},
                )
                if r in weekly_totals:
                    row_data[selected_week] = weekly_totals[r]

                total_shortages_vol = round(sum([
                    float(row_data.get(w, 0.0))
                    for w in ["week 1", "week 2", "week 3", "week 4"]
                ]), 2)
                shortage_amt = round(total_shortages_vol * shortage_rate, 2)

                collected_vol = round(collected_dict.get(r, 0.0), 2)
                gross_earning = round(collected_vol * earning_rate, 2)

                other_ded = round(float(st.session_state.other_deductions.get(r, 0.0)), 2)
                total_book_dr = round(shortage_amt + other_ded, 2)
                net_earning = round(gross_earning - total_book_dr, 2)

                code = active_route_mapping.get(r, "T00")

                row_data["Route"] = row_data.get("Route", r)
                row_data["Total Shortages Vol"] = total_shortages_vol
                row_data["Shortage Amount"] = shortage_amt
                row_data["Collected Volume"] = collected_vol
                row_data["Gross Earning"] = gross_earning
                row_data["Other Deductions"] = other_ded
                row_data["Total Book Dr"] = total_book_dr
                row_data["Net Earning"] = net_earning
                row_data["Code"] = code
                updated_rows.append(row_data)

            master_cols = [
                "Route",
                "week 1",
                "week 2",
                "week 3",
                "week 4",
                "Total Shortages Vol",
                "Shortage Amount",
                "Collected Volume",
                "Gross Earning",
                "Other Deductions",
                "Total Book Dr",
                "Net Earning",
                "Code",
            ]
            df_updated_master = pd.DataFrame(updated_rows)[master_cols]

            # 5. Build Final Totals
            total_row = {
                "Route": "SUM TOTAL",
                "week 1": round(df_updated_master["week 1"].sum(), 2),
                "week 2": round(df_updated_master["week 2"].sum(), 2),
                "week 3": round(df_updated_master["week 3"].sum(), 2),
                "week 4": round(df_updated_master["week 4"].sum(), 2),
                "Total Shortages Vol": round(df_updated_master["Total Shortages Vol"].sum(), 2),
                "Shortage Amount": round(df_updated_master["Shortage Amount"].sum(), 2),
                "Collected Volume": round(df_updated_master["Collected Volume"].sum(), 2),
                "Gross Earning": round(df_updated_master["Gross Earning"].sum(), 2),
                "Other Deductions": round(df_updated_master["Other Deductions"].sum(), 2),
                "Total Book Dr": round(df_updated_master["Total Book Dr"].sum(), 2),
                "Net Earning": round(df_updated_master["Net Earning"].sum(), 2),
                "Code": "",
            }
            df_final_master = pd.concat(
                [df_updated_master, pd.DataFrame([total_row])],
                ignore_index=True,
            )

            # 6. Build Properly Aligned Side-By-Side Positive vs. Negative Sheet
            df_pos_earnings = df_updated_master[df_updated_master["Net Earning"] >= 0][
                ["Route", "Code", "Net Earning"]
            ].reset_index(drop=True)

            df_neg_earnings = df_updated_master[df_updated_master["Net Earning"] < 0][
                ["Route", "Code", "Net Earning"]
            ].reset_index(drop=True)

            pos_tot = round(df_pos_earnings["Net Earning"].sum(), 2)
            neg_tot = round(df_neg_earnings["Net Earning"].sum(), 2)

            max_len = max(len(df_pos_earnings), len(df_neg_earnings), 1)

            df_pos_padded = df_pos_earnings.reindex(range(max_len)).fillna("")
            df_neg_padded = df_neg_earnings.reindex(range(max_len)).fillna("")

            df_net_earnings_sheet = pd.DataFrame({
                "Positive Route": list(df_pos_padded["Route"]) + ["TOTAL POSITIVE"],
                "Positive Code": list(df_pos_padded["Code"]) + [""],
                "Positive Net Earning": list(df_pos_padded["Net Earning"]) + [pos_tot],
                " ": [""] * (max_len + 1),
                "Negative Route": list(df_neg_padded["Route"]) + ["TOTAL NEGATIVE"],
                "Negative Code": list(df_neg_padded["Code"]) + [""],
                "Negative Net Earning": list(df_neg_padded["Net Earning"]) + [neg_tot],
            })

            daily_sheet_name = f"{selected_week.capitalize()} Daily"

            st.success(f"Successfully processed **{selected_week}**!")

            # Executive Summary Metrics
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Gross Earnings", f"KSh {df_updated_master['Gross Earning'].sum():,.2f}")
            m2.metric("Shortages Amount", f"KSh {df_updated_master['Shortage Amount'].sum():,.2f}")
            m3.metric("Total Book Dr (Deductions)", f"KSh {df_updated_master['Total Book Dr'].sum():,.2f}")
            m4.metric("Net Earnings", f"KSh {df_updated_master['Net Earning'].sum():,.2f}")

            tab1, tab2, tab3 = st.tabs([
                "📊 Cumulative Monthly Summary",
                f"📅 {selected_week.capitalize()} Daily Breakdown",
                "📉 Net Earnings Analysis (Positive / Negative)",
            ])
            with tab1:
                st.dataframe(df_final_master, use_container_width=True)
            with tab2:
                st.dataframe(df_daily_grid_final, use_container_width=True)
            with tab3:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("#### 🟢 Routes with Positive Net Earnings")
                    st.dataframe(
                        pd.concat([df_pos_earnings, pd.DataFrame([{
                            "Route": "TOTAL POSITIVE", "Code": "", "Net Earning": pos_tot
                        }])], ignore_index=True),
                        use_container_width=True
                    )
                with c2:
                    st.markdown("#### 🔴 Routes with Negative Net Earnings")
                    st.dataframe(
                        pd.concat([df_neg_earnings, pd.DataFrame([{
                            "Route": "TOTAL NEGATIVE", "Code": "", "Net Earning": neg_tot
                        }])], ignore_index=True),
                        use_container_width=True
                    )

            # 7. Write to Excel Output Safely
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_final_master.to_excel(writer, sheet_name="Monthly Summary", index=False)
                df_net_earnings_sheet.to_excel(writer, sheet_name="Net Earnings Analysis", index=False)
                df_daily_grid_final.to_excel(writer, sheet_name=daily_sheet_name, index=False)

                for s_name, s_df in existing_sheets.items():
                    if s_name not in ["Monthly Summary", "Net Earnings Analysis", daily_sheet_name]:
                        s_df.to_excel(writer, sheet_name=s_name, index=False)

                style_excel_workbook(writer, df_final_master)

            st.download_button(
                label="📥 Download Formatted Master Workbook (.xlsx)",
                data=output.getvalue(),
                file_name=f"Monthly_Route_Summary_Master_{selected_week.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"Error processing workbook: {e}")
