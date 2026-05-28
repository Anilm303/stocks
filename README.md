# Nepal's Stocks Backtest

This folder contains a Python backtest template for trading up to 30 stocks from CSV or Excel data with 14 years of historical data.

## What it does

- Reads price data from a folder of CSV files, a CSV file, or an Excel workbook.
- Uses either a `Signal` column or a confluence strategy built from `SMA_20`, `RSI_14`, and `MACD` when those indicator columns exist.
- Starts with Rs. 2,00,000 by default.
- Writes a combined Excel report plus separate buy-trades and sell-trades Excel files.

## Expected input

Your input can be either:

- A folder like `Data/` containing one CSV per stock.
- A single CSV file.
- An Excel workbook with one or more sheets.

If `Signal` exists, the script uses it directly:

- `1` or `Buy` means buy.
- `-1` or `Sell` means sell.
- `0` or `Hold` means do nothing.

If `Signal` does not exist, the script creates one using a 20-day simple moving average crossover.

## Run

```powershell
C:/Users/Lenovo/AppData/Local/Programs/Python/Python313/python.exe trade_30_stocks.py Data --output trading_report.xlsx
```

## Example

```powershell
C:/Users/Lenovo/AppData/Local/Programs/Python/Python313/python.exe trade_30_stocks.py Data --initial-cash 200000 --max-positions 30
```

## Output Files

- `trading_report.xlsx` - combined report with all sheets
- `trading_report_buy.xlsx` - buy trades workbook with only `BuyTrades`
- `trading_report_sell.xlsx` - sell trades workbook with only `SellTrades`

## Combined Report Sheets

- `Summary`
- `BuyTrades`
- `SellTrades`
- `Trades`
- `EquityCurve`
- `OpenPositions`
- `Signals`

## Notes

- This is an end-of-day style backtest template, not a live trading bot.
- If you want, I can adapt it to your exact Excel layout and strategy rules.