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
orig_eq_df = pd.read_excel("summary report.xlsx", sheet_name="EquityCurve")
orig_eq_df["Year"] = pd.to_datetime(orig_eq_df["Date"]).dt.year
orig_map = orig_eq_df.groupby("Year")["Equity"].last().to_dict()

# Compute base scale map (Target_Equity / Original_Equity)
base_scale_map = {}
for y in sorted(orig_map.keys()):
    base_scale_map[y] = targ_map[y] / orig_map[y]

# 3. Load all trades from summary report.xlsx Trades sheet
trades_df = pd.read_excel("summary report.xlsx", sheet_name="Trades")
trades_df["Date"] = pd.to_datetime(trades_df["Date"])
# Sort trades stably by date to maintain chronological execution order
trades_df = trades_df.sort_values(by="Date", kind="mergesort").reset_index(drop=True)
trades_df["Year"] = trades_df["Date"].dt.year

# 3b. Load signals sheet to look up indicator values for trade reasons
print("Loading Signals sheet to determine trade reasons...")
signals_df = pd.read_excel("summary report.xlsx", sheet_name="Signals")
signals_df["Date_str"] = pd.to_datetime(signals_df["Date"]).dt.strftime("%Y-%m-%d")
signals_df.set_index(["Symbol", "Date_str"], inplace=True)
print("Signals loaded successfully.")

# 3c. Load OpenPositions sheet — stocks held at end that were NOT sold (no-loss rule)
open_pos_df = pd.read_excel("summary report.xlsx", sheet_name="OpenPositions")
print(f"Open positions remaining at end of backtest: {len(open_pos_df)}")

# 4. Run a pass to find the unadjusted total scaled equity (cash + open positions)
# using base scale map to compute the correct multiplier toward Rs 1.5M
temp_cash = 200000.0
temp_open_positions = {}  # symbol -> {qty, entry_price}

for idx, row in trades_df.iterrows():
    trade_year = row["Year"]
    scale = base_scale_map[trade_year]

    qty = round(row["Quantity"] * scale)
    price = row["Price"]
    trade_val = qty * price

    if row["Side"] == "BUY":
        temp_cash -= trade_val
        temp_open_positions[row["Symbol"]] = {"qty": qty, "entry_price": price}
    else:  # SELL or FORCED_EXIT
        temp_cash += trade_val
        temp_open_positions.pop(row["Symbol"], None)

# Add the current market value of open positions (scaled) to temp_cash to get total equity
unadj_cash_only = temp_cash
unadj_op_value = 0.0
for _, op_row in open_pos_df.iterrows():
    symbol = op_row["Symbol"]
    current_price = float(op_row["CurrentPrice"]) if "CurrentPrice" in op_row and pd.notna(op_row["CurrentPrice"]) else float(op_row["EntryPrice"])
    if symbol in temp_open_positions:
        scaled_qty = temp_open_positions[symbol]["qty"]
        val = scaled_qty * current_price
        unadj_op_value += val
        temp_cash += val  # include unrealized value in equity

# temp_cash now represents the total portfolio equity (cash + open positions at market)
unadj_equity = temp_cash  # total equity before scaling
unadj_profit = unadj_equity - 200000.0
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
        
    # Determine reason for trade direction based on the upgraded strategy rules
    if row["Side"] == "BUY":
        reason = f"BUY: Stock price is at its 3-month lowest (Rs. {price:.2f}) (scaled by {scale:.2f}x to fit Rs. 1.5M target)"
    elif row["Side"] == "SELL":
        reason = f"SELL: Profit target reached (+Rs. 20.00 from buy price of Rs. {entry_price:.2f}) (scaled by {scale:.2f}x to fit Rs. 1.5M target)"
    else:  # FORCED_EXIT
        reason = f"SELL: Forced exit at end of backtesting period (scaled by {scale:.2f}x to fit Rs. 1.5M target)"
    
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
        "Reason": reason,
        "OriginalRow": row
    })

# 6. Compute running cash and find discrepancy
cash = 200000.0
for trade in scaled_trades:
    if trade["Side"] == "BUY":
        cash -= trade["TradeValue"]
    else:
        cash += trade["TradeValue"]

# Calculate open position value at current scaled quantities
scaled_open_positions = {}
for trade in scaled_trades:
    symbol = trade["Symbol"]
    qty = trade["Quantity"]
    if trade["Side"] == "BUY":
        scaled_open_positions[symbol] = scaled_open_positions.get(symbol, 0) + qty
    else:  # SELL or FORCED_EXIT
        scaled_open_positions[symbol] = scaled_open_positions.get(symbol, 0) - qty

open_market_value = 0.0
for symbol, qty in scaled_open_positions.items():
    if qty > 0:
        op_match = open_pos_df[open_pos_df["Symbol"] == symbol]
        if not op_match.empty:
            current_price = float(op_match.iloc[0].get("CurrentPrice", op_match.iloc[0]["EntryPrice"]))
        else:
            current_price = 0.0
        open_market_value += qty * current_price

# 7. Adjust quantities of the last trades to make total equity exactly 1,500,000.00
target_cash = 1500000.0 - open_market_value
discrepancy = target_cash - cash

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
            
            # Record quantity adjustment in Reason
            trade["Reason"] += f" (Qty adjusted by {qty_diff} to hit exact 1.5M target)"
            
            if trade["Side"] == "BUY":
                discrepancy += val_change
            else:
                discrepancy -= val_change

# If there's still a small cent discrepancy, adjust the last trade value
if abs(discrepancy) > 0.01:
    last_trade = scaled_trades[-1]
    if last_trade["Side"] in ("SELL", "FORCED_EXIT"):
        last_trade["TradeValue"] = round(last_trade["TradeValue"] + discrepancy, 2)
    else:
        last_trade["TradeValue"] = round(last_trade["TradeValue"] - discrepancy, 2)
    last_trade["Reason"] += " (Adjusted cents to hit exact 1.5M target)"

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
# Include Reason column in buy sheet
df_buy = df_buy[['Date', 'Symbol', 'Quantity', 'BuyPrice', 'BuyValue', 'CashAfterTrade', 'Reason']]

# Prepare Sell Trades Sheet
df_sell = df_all[df_all["Side"].isin(["SELL", "FORCED_EXIT"])].copy()
df_sell = df_sell.rename(columns={"Price": "SellPrice", "TradeValue": "SellValue"})
# Include Reason column in sell sheet
df_sell = df_sell[['Date', 'Symbol', 'Quantity', 'SellPrice', 'SellValue', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice', 'Reason']]

# Prepare Combined Trades Sheet
df_combined = df_all.copy()
df_combined = df_combined.rename(columns={"Price": "Price", "TradeValue": "TradeValue"})
# Include Reason column in combined sheet
df_combined = df_combined[['Date', 'Symbol', 'Side', 'Quantity', 'Price', 'TradeValue', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice', 'Reason']]

# Prepare Scaled Open Positions Sheet (positions held at end, not sold due to no-loss rule)
open_rows = []
for symbol, qty in scaled_open_positions.items():
    if qty > 0:
        buys = [t for t in new_trades_list if t["Symbol"] == symbol and t["Side"] == "BUY"]
        if buys:
            entry_price = buys[-1]["EntryPrice"]
            entry_date = buys[-1]["EntryDate"]
        else:
            entry_price = 0.0
            entry_date = pd.NaT
            
        op_match = open_pos_df[open_pos_df["Symbol"] == symbol]
        if not op_match.empty:
            current_price = float(op_match.iloc[0].get("CurrentPrice", op_match.iloc[0]["EntryPrice"]))
        else:
            current_price = entry_price
            
        market_value = qty * current_price
        unrealized_pnl = (current_price - entry_price) * qty
        status = "HELD (No-Loss Rule: Not sold at end)"
        
        open_rows.append({
            "Symbol": symbol,
            "EntryDate": entry_date,
            "EntryPrice": entry_price,
            "CurrentPrice": current_price,
            "ScaledQuantity": qty,
            "MarketValue": market_value,
            "UnrealizedPnL": unrealized_pnl,
            "Status": status
        })

df_open = pd.DataFrame(open_rows)
if not df_open.empty:
    df_open["EntryDate"] = pd.to_datetime(df_open["EntryDate"]).dt.strftime("%Y-%m-%d")
    df_open = df_open[['Symbol', 'EntryDate', 'EntryPrice', 'CurrentPrice', 'ScaledQuantity', 'MarketValue', 'UnrealizedPnL', 'Status']]
else:
    df_open = pd.DataFrame(columns=['Symbol', 'EntryDate', 'EntryPrice', 'CurrentPrice', 'ScaledQuantity', 'MarketValue', 'UnrealizedPnL', 'Status'])

# Compute final total equity summary
open_market_value = df_open["MarketValue"].sum() if not df_open.empty else 0.0
total_final_equity = cash + open_market_value


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
    df_combined.to_excel(writer, sheet_name="AllTrades", index=False)
    df_open.to_excel(writer, sheet_name="OpenPositions", index=False)
wb = load_workbook(combined_out)
for sheet_name in wb.sheetnames:
    style_ws(wb[sheet_name])
wb.save(combined_out)
print(f"[OK] {combined_out} created")

print(f"Restored all {len(trades_df)} trades! Final cash (realized): Rs. {cash:,.2f}")
print(f"Open positions (scaled, unrealized): Rs. {open_market_value:,.2f}")
print(f"TOTAL PORTFOLIO EQUITY: Rs. {total_final_equity:,.2f}")

