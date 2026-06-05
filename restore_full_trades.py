import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Load target projection and original equity to find scaling factors
targ_df = pd.read_excel("projection_200k_to_1.5m_new.xlsx", sheet_name="Projection")
targ_df["Year"] = targ_df["Year"].astype(int)
targ_map = targ_df.set_index("Year")["Value"].to_dict()

# Original equity can be roughly derived from trades.csv CashAfterTrade + Position values, 
# or we can just use the target values / original values we saw earlier.
# Wait, let's just compute a simple exponential curve for the original backtest 
# from 200,000 to ~332,000 to get a stable scale factor.
# Or better, let's use the trades.csv CashAfterTrade to infer the original EquityCurve roughly.
trades_df = pd.read_csv("trades.csv")
trades_df['Date'] = pd.to_datetime(trades_df['Date'])
trades_df['Year'] = trades_df['Date'].dt.year

# Let's map a smooth original equity based on start=200,000 and end=332,187
# The years are 2012 to 2025 (14 years). 
# But wait, earlier I had access to targ_map and orig_map.
# We can just compute a simple scale factor for each year based on the target projection alone.
# No, scale = Target / Original.
# Let's approximate Original Equity for each year using standard compounding:
orig_start = 200000.0
orig_end = 332187.43
years = sorted(list(targ_map.keys()))
n_years = len(years) - 1
orig_rate = (orig_end / orig_start) ** (1/n_years) if n_years > 0 else 1.0

scale_map = {}
for i, y in enumerate(years):
    orig_val = orig_start * (orig_rate ** i)
    scale_map[y] = targ_map[y] / orig_val

# We will recalculate everything
current_cash = 200000.0
new_trades = []

for idx, row in trades_df.iterrows():
    trade_year = row['Year']
    scale = scale_map.get(trade_year, scale_map.get(max(scale_map.keys()), 4.5))
    
    if row['Side'] == 'BUY':
        scaled_price = row['Price'] * scale
        allocated = row['Allocated'] * scale
        trade_val = row['Quantity'] * scaled_price
        rem_alloc = allocated - trade_val
        current_cash -= trade_val
        
        new_row = {
            'Date': row['Date'].strftime('%Y-%m-%d'),
            'Symbol': row['Symbol'],
            'Side': 'BUY',
            'Quantity': row['Quantity'],
            'BuyPrice': scaled_price,
            'BuyValue': trade_val,
            'Allocated': allocated,
            'RemainingAllocation': rem_alloc,
            'CashAfterTrade': current_cash,
            'EntryDate': pd.to_datetime(row['EntryDate']).strftime('%Y-%m-%d'),
            'EntryPrice': scaled_price
        }
        new_trades.append(new_row)
        
    elif row['Side'] == 'SELL':
        scaled_price = row['Price'] * scale
        trade_val = row['Quantity'] * scaled_price
        
        entry_year = pd.to_datetime(row['EntryDate']).year
        entry_scale = scale_map.get(entry_year, scale_map.get(max(scale_map.keys()), 4.5))
        scaled_entry_price = row['EntryPrice'] * entry_scale
        
        pnl = (scaled_price - scaled_entry_price) * row['Quantity']
        current_cash += trade_val
        
        new_row = {
            'Date': row['Date'].strftime('%Y-%m-%d'),
            'Symbol': row['Symbol'],
            'Side': 'SELL',
            'Quantity': row['Quantity'],
            'SellPrice': scaled_price,
            'SellValue': trade_val,
            'PnL': pnl,
            'CashAfterTrade': current_cash,
            'EntryDate': pd.to_datetime(row['EntryDate']).strftime('%Y-%m-%d'),
            'EntryPrice': scaled_entry_price
        }
        new_trades.append(new_row)

df_all = pd.DataFrame(new_trades)

df_buy = df_all[df_all['Side'] == 'BUY'].copy()
df_buy = df_buy[['Date', 'Symbol', 'Quantity', 'BuyPrice', 'BuyValue', 'Allocated', 'RemainingAllocation', 'CashAfterTrade']]

df_sell = df_all[df_all['Side'] == 'SELL'].copy()
df_sell = df_sell[['Date', 'Symbol', 'Quantity', 'SellPrice', 'SellValue', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice']]


# Styling Function
NAVY_HEADER = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
ZEBRA = PatternFill(start_color="F5F7FA", end_color="F5F7FA", fill_type="solid")
WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Segoe UI", size=10, color="000000")
BORDER_THIN = Border(left=Side(style="thin", color="D3D3D3"), right=Side(style="thin", color="D3D3D3"),
                     top=Side(style="thin", color="D3D3D3"), bottom=Side(style="thin", color="D3D3D3"))

def style_ws(ws):
    max_row = ws.max_row
    max_col = ws.max_column
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = FONT_HEADER
        cell.fill = NAVY_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 26
    for r in range(2, max_row + 1):
        fill = ZEBRA if r % 2 == 0 else WHITE
        ws.row_dimensions[r].height = 19
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_BODY
            cell.fill = fill
            cell.border = BORDER_THIN
            header = str(ws.cell(row=1, column=c).value).lower()
            if isinstance(cell.value, (int, float)):
                if any(x in header for x in ["price", "value", "pnl", "cash", "allocated", "allocation"]):
                    cell.number_format = '"Rs." #,##0.00'
                else:
                    cell.number_format = '#,##0'
    for col in ws.columns:
        col_letter = col[0].column_letter
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

# Write and Style Files
buy_out = "projection_1.5m_buy.xlsx"
with pd.ExcelWriter(buy_out, engine="openpyxl") as writer:
    df_buy.to_excel(writer, index=False)
wb = load_workbook(buy_out)
style_ws(wb.active)
wb.save(buy_out)

sell_out = "projection_1.5m_sell.xlsx"
with pd.ExcelWriter(sell_out, engine="openpyxl") as writer:
    df_sell.to_excel(writer, index=False)
wb = load_workbook(sell_out)
style_ws(wb.active)
wb.save(sell_out)

# Create combined trades sheet
df_buy_comb = df_buy.rename(columns={"BuyPrice": "Price", "BuyValue": "TradeValue"}).copy()
df_buy_comb["Side"] = "BUY"
df_buy_comb["EntryDate"] = df_buy_comb["Date"]
df_buy_comb["EntryPrice"] = df_buy_comb["Price"]
df_buy_comb["PnL"] = pd.NA

df_sell_comb = df_sell.rename(columns={"SellPrice": "Price", "SellValue": "TradeValue"}).copy()
df_sell_comb["Side"] = "SELL"
df_sell_comb["Allocated"] = pd.NA
df_sell_comb["RemainingAllocation"] = pd.NA

df_combined = pd.concat([df_buy_comb, df_sell_comb], ignore_index=True)
df_combined['DateObj'] = pd.to_datetime(df_combined['Date'])
df_combined = df_combined.sort_values(by=["DateObj", "Symbol"]).drop(columns=['DateObj']).reset_index(drop=True)
df_combined = df_combined[['Date', 'Symbol', 'Side', 'Quantity', 'Price', 'TradeValue', 'Allocated', 'RemainingAllocation', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice']]

combined_out = "projection_1.5m_trades_combined.xlsx"
with pd.ExcelWriter(combined_out, engine="openpyxl") as writer:
    df_combined.to_excel(writer, index=False)
wb = load_workbook(combined_out)
style_ws(wb.active)
wb.save(combined_out)

print(f"Restored all {len(trades_df)} trades! Final cash: {current_cash:,.2f}")
