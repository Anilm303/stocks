"""
generate_yearly_sheets.py
=========================
Creates two Excel files:
  - yearly_buy_detail.xlsx    → Buy trades grouped by year (each year: stocks, qty, price, value + yearly total)
  - yearly_sell_detail.xlsx   → Sell trades grouped by year (each year: stocks, qty, price, value, P&L + yearly total)

Format:
  ┌──────────────────────────────────────────────┐
  │   2012                                       │  ← Year header (bold, colored)
  ├──────┬────────┬──────────┬──────────┬────────┤
  │ Date │ Symbol │ Quantity │ BuyPrice │ BuyVal │  ← Column headers
  ├──────┼────────┼──────────┼──────────┼────────┤
  │ ...  │  ...   │   ...    │   ...    │   ...  │  ← Rows
  ├──────┼────────┼──────────┼──────────┼────────┤
  │      │ TOTAL  │   ###    │          │  ###   │  ← Year total row
  └──────┴────────┴──────────┴──────────┴────────┘
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Color Palette ─────────────────────────────────────────────────────────────
YEAR_HEADER_BUY_FILL  = PatternFill("solid", fgColor="1B4F72")   # Deep blue
YEAR_HEADER_SELL_FILL = PatternFill("solid", fgColor="512E5F")   # Deep purple
COL_HEADER_FILL       = PatternFill("solid", fgColor="2E86C1")   # Medium blue (buy)
COL_HEADER_SELL_FILL  = PatternFill("solid", fgColor="7D3C98")   # Medium purple (sell)
TOTAL_FILL            = PatternFill("solid", fgColor="D6EAF8")   # Light blue
TOTAL_SELL_FILL       = PatternFill("solid", fgColor="E8DAEF")   # Light purple
ROW_EVEN_FILL         = PatternFill("solid", fgColor="EBF5FB")
ROW_ODD_FILL          = PatternFill("solid", fgColor="FFFFFF")
PROFIT_FILL           = PatternFill("solid", fgColor="D5F5E3")   # Light green
LOSS_FILL             = PatternFill("solid", fgColor="FADBD8")   # Light red

# ── Fonts ─────────────────────────────────────────────────────────────────────
FONT_YEAR     = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
FONT_COL_HDR  = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_BODY     = Font(name="Segoe UI", size=10, color="000000")
FONT_TOTAL    = Font(name="Segoe UI", size=10, bold=True, color="154360")
FONT_PROFIT   = Font(name="Segoe UI", size=10, bold=True, color="1E8449")
FONT_LOSS     = Font(name="Segoe UI", size=10, bold=True, color="C0392B")

# ── Borders ───────────────────────────────────────────────────────────────────
THIN = Side(style="thin", color="BDC3C7")
MED  = Side(style="medium", color="5D6D7E")

def thin_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def bottom_med_border():
    return Border(left=THIN, right=THIN, top=THIN, bottom=MED)

def top_med_border():
    return Border(left=THIN, right=THIN, top=MED, bottom=THIN)


def write_year_block(ws, year, df_year, row_cursor, columns_config, is_sell=False):
    """
    Write one year's block of data into the worksheet.
    columns_config: list of (header_label, df_col_name, number_format, width)
    Returns the next row cursor after writing this block.
    """
    year_fill  = YEAR_HEADER_SELL_FILL if is_sell else YEAR_HEADER_BUY_FILL
    col_fill   = COL_HEADER_SELL_FILL  if is_sell else COL_HEADER_FILL
    tot_fill   = TOTAL_SELL_FILL       if is_sell else TOTAL_FILL
    num_cols   = len(columns_config)
    end_col    = num_cols

    # ── Year header row ────────────────────────────────────────────────────────
    ws.row_dimensions[row_cursor].height = 24
    for c in range(1, end_col + 1):
        cell = ws.cell(row=row_cursor, column=c)
        cell.fill = year_fill
        cell.border = thin_border()
        if c == 1:
            cell.value = f"  {year}"
            cell.font  = FONT_YEAR
            cell.alignment = Alignment(horizontal="left", vertical="center")
    # Merge year label across all columns
    ws.merge_cells(start_row=row_cursor, start_column=1,
                   end_row=row_cursor, end_column=end_col)
    row_cursor += 1

    # ── Column header row ─────────────────────────────────────────────────────
    ws.row_dimensions[row_cursor].height = 20
    for c_idx, (header, _, _, _) in enumerate(columns_config, start=1):
        cell = ws.cell(row=row_cursor, column=c_idx)
        cell.value     = header
        cell.font      = FONT_COL_HDR
        cell.fill      = col_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = thin_border()
    row_cursor += 1

    # ── Data rows ──────────────────────────────────────────────────────────────
    for i, (_, row_data) in enumerate(df_year.iterrows()):
        ws.row_dimensions[row_cursor].height = 18
        row_fill = ROW_EVEN_FILL if i % 2 == 0 else ROW_ODD_FILL

        for c_idx, (_, col_name, num_fmt, _) in enumerate(columns_config, start=1):
            cell = ws.cell(row=row_cursor, column=c_idx)
            val  = row_data.get(col_name, "")

            # Clean up NaN
            if pd.isna(val):
                val = ""

            cell.value  = val
            cell.border = thin_border()

            if isinstance(val, (int, float)) and val != "":
                cell.number_format = num_fmt
                cell.alignment = Alignment(horizontal="right", vertical="center")
                # Color P&L column
                if col_name == "PnL":
                    if val >= 0:
                        cell.fill = PROFIT_FILL
                        cell.font = FONT_PROFIT
                    else:
                        cell.fill = LOSS_FILL
                        cell.font = FONT_LOSS
                else:
                    cell.fill = row_fill
                    cell.font = FONT_BODY
            else:
                cell.fill      = row_fill
                cell.font      = FONT_BODY
                cell.alignment = Alignment(horizontal="center" if c_idx <= 2 else "left",
                                           vertical="center")
        row_cursor += 1

    # ── Yearly Total row ───────────────────────────────────────────────────────
    ws.row_dimensions[row_cursor].height = 20
    for c_idx, (_, col_name, num_fmt, _) in enumerate(columns_config, start=1):
        cell = ws.cell(row=row_cursor, column=c_idx)
        cell.fill   = tot_fill
        cell.border = bottom_med_border()

        if c_idx == 1:
            cell.value     = f"{year} Total"
            cell.font      = FONT_TOTAL
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif col_name in ("Quantity", "BuyValue", "SellValue", "PnL"):
            total_val         = df_year[col_name].sum()
            cell.value        = total_val
            cell.number_format = num_fmt
            cell.font         = FONT_TOTAL
            cell.alignment    = Alignment(horizontal="right", vertical="center")
        else:
            cell.value = ""
    row_cursor += 1

    # Spacer row
    ws.row_dimensions[row_cursor].height = 8
    row_cursor += 1

    return row_cursor


def build_sheet(ws, df, trade_type="BUY"):
    """Build a complete year-by-year sheet for buy or sell trades."""
    is_sell = (trade_type == "SELL")

    # Title row
    title_text = ("YEARLY SELL TRADES - Nepal Stock Market Backtest (Rs. 200,000 -> Rs. 1,500,000)"
                  if is_sell else
                  "YEARLY BUY TRADES - Nepal Stock Market Backtest (Rs. 200,000 -> Rs. 1,500,000)")
    ws["A1"] = title_text
    ws["A1"].font = Font(name="Segoe UI", size=14, bold=True,
                         color="512E5F" if is_sell else "1B4F72")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    # Subtitle row
    ws["A2"] = "Har barsa kati stocks kine/beche — date, symbol, qty, price ra total value sahit"
    ws["A2"].font = Font(name="Segoe UI", size=9, italic=True, color="5D6D7E")
    ws.row_dimensions[2].height = 15

    ws.row_dimensions[3].height = 10  # spacer

    row_cursor = 4

    if is_sell:
        columns_config = [
            # (Header Label,  df column name,   number_format,    approx width)
            ("Date",          "Date",            "@",               14),
            ("Stock",         "Symbol",          "@",               10),
            ("Qty Sold",      "Quantity",        "#,##0",           10),
            ("Sell Price",    "SellPrice",       '"Rs."#,##0.00',   13),
            ("Sell Value",    "SellValue",       '"Rs."#,##0.00',   14),
            ("Buy Price",     "EntryPrice",      '"Rs."#,##0.00',   13),
            ("Buy Date",      "EntryDate",       "@",               14),
            ("P&L",           "PnL",             '"Rs."#,##0.00',   14),
            ("Cash After",    "CashAfterTrade",  '"Rs."#,##0.00',   15),
        ]
    else:
        columns_config = [
            ("Date",          "Date",            "@",               14),
            ("Stock",         "Symbol",          "@",               10),
            ("Qty Bought",    "Quantity",        "#,##0",           10),
            ("Buy Price",     "BuyPrice",        '"Rs."#,##0.00',   13),
            ("Buy Value",     "BuyValue",        '"Rs."#,##0.00',   14),
            ("Allocated",     "Allocated",       '"Rs."#,##0.00',   13),
            ("Cash After",    "CashAfterTrade",  '"Rs."#,##0.00',   15),
        ]

    # Set column widths upfront
    for c_idx, (_, _, _, width) in enumerate(columns_config, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = width

    # Loop over each year
    years = sorted(df["Year"].dropna().unique())
    for year in years:
        df_year = df[df["Year"] == year].copy()
        row_cursor = write_year_block(ws, int(year), df_year,
                                      row_cursor, columns_config, is_sell=is_sell)

    # ── Grand Total row ────────────────────────────────────────────────────────
    ws.row_dimensions[row_cursor].height = 22
    grand_qty = df["Quantity"].sum()
    value_col  = "SellValue" if is_sell else "BuyValue"
    grand_val  = df[value_col].sum()
    grand_pnl  = df["PnL"].sum() if is_sell else None

    grand_fill = PatternFill("solid", fgColor="1B2631")
    grand_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

    num_cols = len(columns_config)
    for c_idx in range(1, num_cols + 1):
        cell = ws.cell(row=row_cursor, column=c_idx)
        cell.fill   = grand_fill
        cell.font   = grand_font
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="right", vertical="center")

    ws.cell(row=row_cursor, column=1).value     = "  GRAND TOTAL"
    ws.cell(row=row_cursor, column=1).alignment = Alignment(horizontal="left", vertical="center")

    # Qty column (index 3)
    ws.cell(row=row_cursor, column=3).value        = grand_qty
    ws.cell(row=row_cursor, column=3).number_format = "#,##0"

    # Value column (index 5)
    ws.cell(row=row_cursor, column=5).value        = grand_val
    ws.cell(row=row_cursor, column=5).number_format = '"Rs."#,##0.00'

    # PnL column for sell (index 8)
    if is_sell and grand_pnl is not None:
        ws.cell(row=row_cursor, column=8).value        = grand_pnl
        ws.cell(row=row_cursor, column=8).number_format = '"Rs."#,##0.00'
        ws.cell(row=row_cursor, column=8).font         = Font(
            name="Segoe UI", size=11, bold=True,
            color="00FF88" if grand_pnl >= 0 else "FF6B6B"
        )

    # Freeze panes so headers stay visible
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = True


def main():
    print("Reading backtest trade data...")
    df_buy  = pd.read_excel("trading_report.xlsx", "BuyTrades")
    df_sell = pd.read_excel("trading_report.xlsx", "SellTrades")

    df_buy["Year"]  = pd.to_datetime(df_buy["Date"]).dt.year
    df_sell["Year"] = pd.to_datetime(df_sell["Date"]).dt.year

    # ── BUY Excel file ─────────────────────────────────────────────────────────
    print("Writing yearly_buy_detail.xlsx ...")
    wb_buy = Workbook()
    ws_buy = wb_buy.active
    ws_buy.title = "Yearly Buy Detail"
    build_sheet(ws_buy, df_buy, trade_type="BUY")
    wb_buy.save("yearly_buy_detail.xlsx")
    print("  [OK] yearly_buy_detail.xlsx saved")

    # ── SELL Excel file ────────────────────────────────────────────────────────
    print("Writing yearly_sell_detail.xlsx ...")
    wb_sell = Workbook()
    ws_sell = wb_sell.active
    ws_sell.title = "Yearly Sell Detail"
    build_sheet(ws_sell, df_sell, trade_type="SELL")
    wb_sell.save("yearly_sell_detail.xlsx")
    print("  ✓ yearly_sell_detail.xlsx saved")

    # ── Summary print ─────────────────────────────────────────────────────────
    print("\n" + "="*60)
    print("YEARLY BUY SUMMARY")
    print("="*60)
    summary_buy = (df_buy.groupby("Year")
                   .agg(Stocks_Bought=("Symbol", "count"),
                        Total_Qty=("Quantity", "sum"),
                        Total_Buy_Value=("BuyValue", "sum"))
                   .reset_index())
    print(summary_buy.to_string(index=False))

    print("\n" + "="*60)
    print("YEARLY SELL SUMMARY")
    print("="*60)
    summary_sell = (df_sell.groupby("Year")
                    .agg(Stocks_Sold=("Symbol", "count"),
                         Total_Qty=("Quantity", "sum"),
                         Total_Sell_Value=("SellValue", "sum"),
                         Total_PnL=("PnL", "sum"))
                    .reset_index())
    print(summary_sell.to_string(index=False))

    print("\nDone! Files created:")
    print("  → yearly_buy_detail.xlsx")
    print("  → yearly_sell_detail.xlsx")


if __name__ == "__main__":
    main()
