import pandas as pd

# Load week1.xlsx using the dual-mode parser logic to verify row & column total calculations
xls = pd.ExcelFile('week1.xlsx')
df_raw = pd.read_excel(xls, sheet_name=0)

day_cols_info = []
header_has_days = any(str(c).strip().isdigit() for c in df_raw.columns)

if header_has_days:
    for c_idx, col_name in enumerate(df_raw.columns):
        c_str = str(col_name).strip()
        if c_str.isdigit():
            day_cols_info.append((f"Day {int(c_str)}", c_idx - 1, c_idx))
    start_row = 1
else:
    for c_idx in range(len(df_raw.columns)):
        val = df_raw.iloc[0, c_idx]
        if pd.notna(val) and str(val).strip().isdigit():
            day_cols_info.append((f"Day {int(str(val).strip())}", c_idx - 1, c_idx))
    start_row = 2

route_records = {}
for day_label, route_col, val_col in day_cols_info:
    for r in range(start_row, len(df_raw)):
        r_name = df_raw.iloc[r, route_col]
        r_var = df_raw.iloc[r, val_col]
        if pd.notna(r_name) and str(r_name).strip() not in ["", "0", "Route", "nan", "NaN"]:
            r_clean = str(r_name).strip().upper()
            if r_clean not in route_records:
                route_records[r_clean] = {}
            v_num = pd.to_numeric(r_var, errors='coerce') if pd.notna(r_var) else 0.0
            route_records[r_clean][day_label] = float(v_num) if pd.notna(v_num) else 0.0

df_daily_grid = pd.DataFrame.from_dict(route_records, orient='index').fillna(0)
df_daily_grid.index.name = "Route"
df_daily_grid = df_daily_grid.reset_index()
day_cols_sorted = sorted([c for c in df_daily_grid.columns if c != "Route"], key=lambda x: int(x.replace("Day ", "")))

# Add Row Total ('Total')
df_daily_grid["Total"] = df_daily_grid[day_cols_sorted].sum(axis=1)

# Add Column Total ('SUM TOTAL')
daily_totals = {"Route": "SUM TOTAL"}
for col in day_cols_sorted + ["Total"]:
    daily_totals[col] = df_daily_grid[col].sum()

df_daily_grid_final = pd.concat([df_daily_grid, pd.DataFrame([daily_totals])], ignore_index=True)

print("Daily Breakdown grid with Row & Column totals:")
print(df_daily_grid_final.tail(5))
