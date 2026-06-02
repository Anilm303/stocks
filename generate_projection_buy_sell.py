# generate_projection_buy_sell.py
"""
Generate projected BUY and SELL Excel sheets from the projection workbook.
The source file `projection_200k_to_1.5m_new.xlsx` contains a sheet `Projection`
with columns: Year, Date, Value (projected portfolio value at year end).
We compute the incremental change between consecutive years as the BUY amount for that year.
Since the projection is a monotonic growth, SELL amounts are zero.
The script creates two workbooks:
  - projection_200k_to_1.5m_buy.xlsx  (Year, Date, BuyValue)
  - projection_200k_to_1.5m_sell.xlsx (Year, Date, SellValue)
Both files are styled with premium colors and number formats.
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Styling helpers ────────────────────────────────────────────────────────
FILL_HEADER = PatternFill(fill_type="solid", fgColor="1B4F72")  # dark blue
FILL_TOTAL  = PatternFill(fill_type="solid", fgColor="D6EAF8")  # light blue
FILL_EVEN   = PatternFill(fill_type="solid", fgColor="EBF5FB")
FILL_ODD    = PatternFill(fill_type="solid", fgColor="FFFFFF")

FONT_TITLE   = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
FONT_HEADER  = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_BODY    = Font(name="Segoe UI", size=10, color="000000")
FONT_TOTAL   = Font(name="Segoe UI", size=10, bold=True, color="154360")

THIN = Side(style="thin", color="BDC3C7")
MED  = Side(style="medium", color="5D6D7E")

def thin_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def bottom_med_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=MED)

def top_med_border():
    return Border(left=THIN, right=THIN, top=MED, bottom=THIN)

def style_sheet(ws, title_text):
    ws.title = title_text
    ws["A1"] = title_text
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = FILL_HEADER
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    # column widths will be set later
    ws.sheet_view.showGridLines = True

def write_data(ws, df, is_sell=False):
    # Header row (row 3)
    start_row = 3
    ws.row_dimensions[start_row].height = 20
    headers = ["Year", "Date", "Sell Value" if is_sell else "Buy Value"]
    for col_idx, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx)
        cell.value = hdr
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border()
    # Data rows
    for i, row in df.iterrows():
        r = start_row + 1 + i
        ws.row_dimensions[r].height = 18
        fill = FILL_EVEN if i % 2 == 0 else FILL_ODD
        ws.cell(row=r, column=1, value=row["Year"]).font = FONT_BODY
        ws.cell(row=r, column=2, value=row["Date"]).font = FONT_BODY
        val_cell = ws.cell(row=r, column=3, value=row["Amount"])
        val_cell.number_format = '"Rs."#,##0.00'
        val_cell.font = FONT_BODY
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.border = thin_border()
            if c != 3:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="right", vertical="center")
    # Total row
    total_row = start_row + 1 + len(df)
    ws.row_dimensions[total_row].height = 20
    for c in range(1, 4):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = FILL_TOTAL
        cell.border = bottom_med_border()
        cell.font = FONT_TOTAL
    ws.cell(row=total_row, column=1, value="TOTAL").font = FONT_TOTAL
    ws.cell(row=total_row, column=2, value="").value = ""
    total_val = df["Amount"].sum()
    tot_cell = ws.cell(row=total_row, column=3, value=total_val)
    tot_cell.number_format = '"Rs."#,##0.00'
    tot_cell.alignment = Alignment(horizontal="right", vertical="center")
    # Auto‑fit columns
    for col_idx in range(1, 4):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for cell in ws[col_letter]:
            if cell.value:
                l = len(str(cell.value))
                if l > max_len:
                    max_len = l
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

def main():
    src_file = "projection_200k_to_1.5m_new.xlsx"
    xl = pd.ExcelFile(src_file)
    df = pd.read_excel(xl, "Projection")
    # Ensure proper types
    df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype(int)
    df["Date"] = pd.to_datetime(df["Date"]).dt.date
    # Compute incremental change (Buy amount) per year
    df = df.sort_values("Year")
    df["PrevValue"] = df["Value"].shift(1).fillna(0)
    df["BuyAmount"] = df["Value"] - df["PrevValue"]
    # For first year, the whole value is considered as buy (initial capital)
    # Sell amount is zero (projection assumes growth, not sell‑offs)
    buy_df = df[["Year", "Date", "BuyAmount"]].rename(columns={"BuyAmount": "Amount"})
    sell_df = df[["Year", "Date"]].copy()
    sell_df["Amount"] = 0.0
    # Write BUY file
    wb_buy = Workbook()
    ws_buy = wb_buy.active
    style_sheet(ws_buy, "Projected BUY Trades (200k → 1.5M)")
    write_data(ws_buy, buy_df, is_sell=False)
    wb_buy.save("projection_200k_to_1.5m_buy.xlsx")
    print("[OK] projection_200k_to_1.5m_buy.xlsx created")
    # Write SELL file
    wb_sell = Workbook()
    ws_sell = wb_sell.active
    style_sheet(ws_sell, "Projected SELL Trades (200k → 1.5M)")
    write_data(ws_sell, sell_df, is_sell=True)
    wb_sell.save("projection_200k_to_1.5m_sell.xlsx")
    print("[OK] projection_200k_to_1.5m_sell.xlsx created")
    # Write COMBINED BUY/SELL file
    wb_comb = Workbook()
    ws_comb = wb_comb.active
    style_sheet(ws_comb, "Projected BUY & SELL (200k → 1.5M)")
    # Merge buy and sell data on Year and Date
    combined_df = pd.merge(buy_df, sell_df, on=["Year", "Date"], how="inner")
    # Rename for clarity
    combined_df = combined_df.rename(columns={"Amount_x": "BuyAmount", "Amount_y": "SellAmount"})
    # Write header row (will be added by write_data)
    # Prepare dataframe for write_data: include Year, Date, BuyAmount, SellAmount
    write_df = combined_df[["Year", "Date", "BuyAmount", "SellAmount"]]
    # Use write_data with is_sell=False (format both as currency)
    write_data(ws_comb, write_df, is_sell=False)
    # Adjust header names manually after writing
    header_row = 3  # after title and subtitle
    ws_comb.cell(row=header_row, column=3, value="Buy Value").font = FONT_HEADER
    ws_comb.cell(row=header_row, column=4, value="Sell Value").font = FONT_HEADER
    # Adjust number formats for the two amount columns
    for r in range(header_row + 1, ws_comb.max_row + 1):
        ws_comb.cell(row=r, column=3).number_format = '"Rs."#,##0.00'
        ws_comb.cell(row=r, column=4).number_format = '"Rs."#,##0.00'
    # Add total row
    total_row = ws_comb.max_row + 1
    ws_comb.cell(row=total_row, column=1, value="TOTAL").font = FONT_TOTAL
    ws_comb.cell(row=total_row, column=1).alignment = Alignment(horizontal="left")
    ws_comb.cell(row=total_row, column=3, value=combined_df["BuyAmount"].sum()).number_format = '"Rs."#,##0.00'
    ws_comb.cell(row=total_row, column=4, value=combined_df["SellAmount"].sum()).number_format = '"Rs."#,##0.00'
    ws_comb.cell(row=total_row, column=3).font = FONT_TOTAL
    ws_comb.cell(row=total_row, column=4).font = FONT_TOTAL
    wb_comb.save("projection_200k_to_1.5m_buy_sell_combined.xlsx")
    print("[OK] projection_200k_to_1.5m_buy_sell_combined.xlsx created")
    wb_buy = Workbook()
    ws_buy = wb_buy.active
    style_sheet(ws_buy, "Projected BUY Trades (200k → 1.5M)")
    write_data(ws_buy, buy_df, is_sell=False)
    wb_buy.save("projection_200k_to_1.5m_buy.xlsx")
    print("✓ projection_200k_to_1.5m_buy.xlsx created")
    # Write SELL file
    wb_sell = Workbook()
    ws_sell = wb_sell.active
    style_sheet(ws_sell, "Projected SELL Trades (200k → 1.5M)")
    write_data(ws_sell, sell_df, is_sell=True)
    wb_sell.save("projection_200k_to_1.5m_sell.xlsx")
    print("✓ projection_200k_to_1.5m_sell.xlsx created")

if __name__ == "__main__":
    main()
