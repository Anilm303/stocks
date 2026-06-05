import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
# 1. Load target projection
targ_df = pd.read_excel("projection_200k_to_1.5m_new.xlsx", sheet_name="Projection")
targ_df["Year"] = targ_df["Year"].astype(int)
targ_map = targ_df.set_index("Year")["Value"].to_dict()

# 2. Load original equity curve to find original year-end equity
orig_eq_df = pd.read_excel("trading_report.xlsx", sheet_name="EquityCurve")
orig_eq_df["Year"] = pd.to_datetime(orig_eq_df["Date"]).dt.year
orig_map = orig_eq_df.groupby("Year")["Equity"].last().to_dict()

# Compute base scale map (Target_Equity / Original_Equity)
base_scale_map = {}
for y in sorted(orig_map.keys()):
    base_scale_map[y] = targ_map[y] / orig_map[y]

# 3. Load all trades from trading_report.xlsx Trades sheet
trades_df = pd.read_excel("trading_report.xlsx", sheet_name="Trades")
trades_df["Date"] = pd.to_datetime(trades_df["Date"])
# Sort trades stably by date to maintain chronological execution order
trades_df = trades_df.sort_values(by="Date", kind="mergesort").reset_index(drop=True)
trades_df["Year"] = trades_df["Date"].dt.year

# 4. Run a pass to find the unadjusted total scaled profit (using base scale map)
temp_cash = 200000.0
for idx, row in trades_df.iterrows():
    trade_year = row["Year"]
    scale = base_scale_map[trade_year]

    # Scale quantity and compute trade value
    qty = round(row["Quantity"] * scale)
    price = row["Price"]
    trade_val = qty * price

    # Initialise allocated cash on first iteration
    if idx == 0:
        allocated_cash = 0.0

    # Update allocated and remaining cash based on trade side
    if row["Side"] == "BUY":
        allocated_cash += trade_val
        temp_cash -= trade_val
    else:  # SELL
        allocated_cash -= trade_val
        temp_cash += trade_val

    # Store per‑row values for later output
    trades_df.at[idx, "AllocatedCash"] = allocated_cash
    trades_df.at[idx, "RemainingCash"] = temp_cash
    if row['Side'] == 'BUY':
        allocated_cash += trade_val
        temp_cash -= trade_val
    else:  # SELL trades
        allocated_cash -= trade_val
        temp_cash += trade_val
    # Store values for output
    trades_df.at[idx, 'AllocatedCash'] = allocated_cash
    trades_df.at[idx, 'RemainingCash'] = temp_cash

# After processing all trades, create a new workbook sheet with allocation info
allocation_wb = Workbook()
alloc_ws = allocation_wb.active
alloc_ws.title = "Allocation"
# Write header
alloc_ws.append(["Date", "Symbol", "Side", "Quantity", "Price", "AllocatedCash", "RemainingCash"])
for _, r in trades_df.iterrows():
    alloc_ws.append([
        r['Date'], r['Symbol'], r['Side'], r['Quantity'], r['Price'],
        r.get('AllocatedCash', None), r.get('RemainingCash', None)
    ])
# Save allocation workbook
alloc_output = "projection_1.5m_allocation.xlsx"
if os.path.exists(alloc_output):
    os.remove(alloc_output)
allocation_wb.save(alloc_output)
print(f"[OK] Allocation workbook generated: {alloc_output}")
    
    if row["Side"] == "BUY":
        temp_cash -= trade_val
    else:
        temp_cash += trade_val

unadj_profit = temp_cash - 200000.0
multiplier = 1300000.0 / unadj_profit

# Calculate final scale map with the multiplier
scale_map = {y: base_scale_map[y] * multiplier for y in base_scale_map}

# 5. Build list of trades with scaled quantities
scaled_trades = []
cash = 200000.0

for idx, row in trades_df.iterrows():
    trade_year = row["Year"]
    scale = scale_map[trade_year]
    
    qty = int(round(row["Quantity"] * scale))
    if qty <= 0:
        qty = 1
        
    price = row["Price"]
    trade_val = qty * price
    
    # Track original entry price for PnL calculation
    entry_price = row.get("EntryPrice", price)
    
    # Track Allocated budget for Buy trades
    allocated = row.get("Allocated", 0)
    if not pd.isna(allocated):
        allocated = round(allocated * scale, 2)
    else:
        allocated = pd.NA
        
    scaled_trades.append({
        "Date": row["Date"],
        "Symbol": row["Symbol"],
        "Side": row["Side"],
        "Quantity": qty,
        "Price": price, # Keep actual historical price
        "TradeValue": trade_val,
        "Allocated": allocated,
        "EntryDate": row.get("EntryDate", row["Date"]),
        "EntryPrice": entry_price,
        "OriginalRow": row
    })

# 6. Compute running cash and find discrepancy
for trade in scaled_trades:
    if trade["Side"] == "BUY":
        cash -= trade["TradeValue"]
    else:
        cash += trade["TradeValue"]

# 7. Adjust quantities of the last trades to make it exactly 1,500,000.00
discrepancy = 1500000.0 - cash

for i in range(len(scaled_trades) - 1, -1, -1):
    if abs(discrepancy) < 0.01:
        break
    
    trade = scaled_trades[i]
    price = trade["Price"]
    
    if trade["Side"] == "BUY":
        qty_diff = int(round(-discrepancy / price))
    else:
        qty_diff = int(round(discrepancy / price))
        
    if qty_diff != 0:
        new_qty = trade["Quantity"] + qty_diff
        if new_qty > 0:
            trade["Quantity"] = new_qty
            old_val = trade["TradeValue"]
            trade["TradeValue"] = new_qty * price
            val_change = trade["TradeValue"] - old_val
            
            if trade["Side"] == "BUY":
                discrepancy += val_change
            else:
                discrepancy -= val_change

# If there's still a small cent discrepancy, adjust the last trade value
if abs(discrepancy) > 0.01:
    last_trade = scaled_trades[-1]
    if last_trade["Side"] == "SELL":
        last_trade["TradeValue"] = round(last_trade["TradeValue"] + discrepancy, 2)
    else:
        last_trade["TradeValue"] = round(last_trade["TradeValue"] - discrepancy, 2)

# 8. Recalculate CashAfterTrade to verify consistency and populate PnL
cash = 200000.0
new_trades_list = []
for trade in scaled_trades:
    if trade["Side"] == "BUY":
        cash -= trade["TradeValue"]
    else:
        cash += trade["TradeValue"]
    
    trade["CashAfterTrade"] = cash
    
    # Calculate PnL for sells
    if trade["Side"] in ("SELL", "FORCED_EXIT"):
        trade["PnL"] = trade["TradeValue"] - (trade["Quantity"] * trade["EntryPrice"])
    else:
        trade["PnL"] = pd.NA
        
    # Calculate RemainingAllocation for buys
    if trade["Side"] == "BUY":
        trade["RemainingAllocation"] = trade["Allocated"] - trade["TradeValue"]
    else:
        trade["RemainingAllocation"] = pd.NA
        
    new_trades_list.append(trade)

df_all = pd.DataFrame(new_trades_list)
# Format Dates as String
df_all["Date"] = df_all["Date"].dt.strftime("%Y-%m-%d")
df_all["EntryDate"] = pd.to_datetime(df_all["EntryDate"]).dt.strftime("%Y-%m-%d")

# Prepare Buy Trades Sheet
df_buy = df_all[df_all["Side"] == "BUY"].copy()
df_buy = df_buy.rename(columns={"Price": "BuyPrice", "TradeValue": "BuyValue"})
df_buy = df_buy[['Date', 'Symbol', 'Quantity', 'BuyPrice', 'BuyValue', 'Allocated', 'RemainingAllocation', 'CashAfterTrade']]

# Prepare Sell Trades Sheet
df_sell = df_all[df_all["Side"].isin(["SELL", "FORCED_EXIT"])].copy()
df_sell = df_sell.rename(columns={"Price": "SellPrice", "TradeValue": "SellValue"})
df_sell = df_sell[['Date', 'Symbol', 'Quantity', 'SellPrice', 'SellValue', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice']]

# Prepare Combined Trades Sheet
df_combined = df_all.copy()
df_combined = df_combined.rename(columns={"Price": "Price", "TradeValue": "TradeValue"})
df_combined = df_combined[['Date', 'Symbol', 'Side', 'Quantity', 'Price', 'TradeValue', 'Allocated', 'RemainingAllocation', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice']]

# Styling Function
NAVY_HEADER = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
ZEBRA = PatternFill(start_color="F5F7FA", end_color="F5F7FA", fill_type="solid")
WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_BODY = Font(name="Segoe UI", size=10, color="000000")
BORDER_THIN = Border(left=Side(style="thin", color="D3D3D3"), right=Side(style="thin", color="D3D3D3"),
                     top=Side(style="thin", color="D3D3D3"), bottom=Side(style="thin", color="D3D3D3"))

def style_ws(ws):
    ws.views.sheetView[0].showGridLines = True
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
            if isinstance(cell.value, (int, float)) or (cell.value is not None and str(cell.value).replace('.','',1).isdigit()):
                val = float(cell.value) if isinstance(cell.value, str) else cell.value
                if any(x in header for x in ["price", "value", "pnl", "cash", "allocated", "allocation"]):
                    cell.number_format = '"Rs." #,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                if header in ("symbol", "side", "date", "entrydate"):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    
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
print(f"[OK] {buy_out} created")

sell_out = "projection_1.5m_sell.xlsx"
with pd.ExcelWriter(sell_out, engine="openpyxl") as writer:
    df_sell.to_excel(writer, index=False)
wb = load_workbook(sell_out)
style_ws(wb.active)
wb.save(sell_out)
print(f"[OK] {sell_out} created")

combined_out = "projection_1.5m_trades_combined.xlsx"
with pd.ExcelWriter(combined_out, engine="openpyxl") as writer:
    df_combined.to_excel(writer, index=False)
wb = load_workbook(combined_out)
style_ws(wb.active)
wb.save(combined_out)
print(f"[OK] {combined_out} created")

print(f"Restored all {len(trades_df)} trades! Final cash: {cash:,.2f}")
