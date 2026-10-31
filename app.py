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
                    if pd.notna(val) and str(val).strip().isdigit()
