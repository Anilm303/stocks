import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os

# Define Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_DIR = os.path.join(BASE_DIR, "Excel_Files")

def get_path(filename):
    return os.path.join(EXCEL_DIR, filename)

# Load original trades
buy_df = pd.read_excel(get_path("trading_report_buy.xlsx"))
sell_df = pd.read_excel(get_path("trading_report_sell.xlsx"))

# Load target projection and original equity to find scaling factors
targ_df = pd.read_excel(get_path("projection_200k_to_1.5m_new.xlsx"), sheet_name="Projection")
targ_df["Year"] = targ_df["Year"].astype(int)
targ_map = targ_df.set_index("Year")["Value"].to_dict()

orig_eq = pd.read_excel(get_path("trading_report.xlsx"), sheet_name="EquityCurve")
orig_eq["Year"] = pd.to_datetime(orig_eq["Date"]).dt.year
orig_map = orig_eq.groupby("Year")["Equity"].last().to_dict()

scale_map = {}
for y in orig_map:
    scale_map[y] = targ_map.get(y, orig_map[y]) / orig_map[y]

# Combine trades into a single timeline to recalculate cash
buy_df['Type'] = 'BUY'
sell_df['Type'] = 'SELL'
all_trades = pd.concat([buy_df, sell_df]).sort_values(by="Date").reset_index(drop=True)

# We will recalculate everything
current_cash = 200000.0
new_trades = []

for idx, row in all_trades.iterrows():
    trade_year = pd.to_datetime(row['Date']).year
    scale = scale_map.get(trade_year, 1.0)
    
    if row['Type'] == 'BUY':
        scaled_price = row['BuyPrice'] * scale
        allocated = row['Allocated'] * scale

        # Enforce multiple of 10 and minimum 20 rule
        qty = (row['Quantity'] // 10) * 10
        if qty < 20: qty = 20

        trade_val = qty * scaled_price
        rem_alloc = allocated - trade_val
        current_cash -= trade_val
        
        new_row = {
            'Date': row['Date'],
            'Symbol': row['Symbol'],
            'Side': 'BUY',
            'Quantity': qty,
            'Price': scaled_price,
            'TradeValue': trade_val,
            'Allocated': allocated,
            'RemainingAllocation': rem_alloc,
            'CashAfterTrade': current_cash,
            'EntryDate': row['EntryDate'],
            'EntryPrice': scaled_price
        }
        new_trades.append(new_row)
        
    elif row['Type'] == 'SELL':
        # For sell, price is scaled by current year
        scaled_price = row['SellPrice'] * scale

        # Enforce multiple of 10 and minimum 20 rule (must match the BUY)
        qty = (row['Quantity'] // 10) * 10
        if qty < 20: qty = 20

        trade_val = qty * scaled_price
        
        # Entry price must be scaled by the EntryDate year to match the BUY record
        entry_year = pd.to_datetime(row['EntryDate']).year
        entry_scale = scale_map.get(entry_year, 1.0)
        scaled_entry_price = row['EntryPrice'] * entry_scale
        
        pnl = (scaled_price - scaled_entry_price) * row['Quantity']
        current_cash += trade_val
        
        new_row = {
            'Date': row['Date'],
            'Symbol': row['Symbol'],
            'Side': 'SELL',
            'Quantity': row['Quantity'],
            'Price': scaled_price,
            'TradeValue': trade_val,
            'PnL': pnl,
            'CashAfterTrade': current_cash,
            'EntryDate': row['EntryDate'],
            'EntryPrice': scaled_entry_price
        }
        new_trades.append(new_row)

df_all = pd.DataFrame(new_trades)

# Split back to buy and sell
df_buy = df_all[df_all['Side'] == 'BUY'].copy()
df_buy = df_buy.rename(columns={'Price': 'BuyPrice', 'TradeValue': 'BuyValue'})
df_buy = df_buy[['Date', 'Symbol', 'Quantity', 'BuyPrice', 'BuyValue', 'Allocated', 'RemainingAllocation', 'CashAfterTrade', 'EntryDate', 'EntryPrice']]

df_sell = df_all[df_all['Side'] == 'SELL'].copy()
df_sell = df_sell.rename(columns={'Price': 'SellPrice', 'TradeValue': 'SellValue'})
df_sell = df_sell[['Date', 'Symbol', 'Quantity', 'SellPrice', 'SellValue', 'PnL', 'CashAfterTrade', 'EntryDate', 'EntryPrice']]

print(f"Final cash after all trades: {current_cash:,.2f}")

# Write to Excel
output_file = get_path("projection_1.5m_buy_sell.xlsx")
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_buy.to_excel(writer, sheet_name="Buy Trades", index=False)
    df_sell.to_excel(writer, sheet_name="Sell Trades", index=False)

# Apply Indian/Nepali Number Format (Lakhs/Crores)
workbook = load_workbook(output_file)
financial_columns = {"BuyPrice", "BuyValue", "SellPrice", "SellValue", "PnL", "CashAfterTrade", "EntryPrice", "Allocated", "RemainingAllocation"}
# Format as 00,00,000 style using Indian Locale
number_format = '[$-en-IN]#,##,##,##0.00'

for worksheet in workbook.worksheets:
    header_map = {cell.value: cell.column for cell in worksheet[1] if cell.value in financial_columns}
    for column_index in header_map.values():
        column_letter = get_column_letter(column_index)
        for row_index in range(2, worksheet.max_row + 1):
            worksheet[f"{column_letter}{row_index}"].number_format = number_format

workbook.save(output_file)

print(f"Done writing to {output_file} with Indian number formatting")
