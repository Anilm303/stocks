# Nepal's Stocks Backtest

This folder contains a Python backtest template for trading up to 30 stocks from an Excel workbook with 14 years of historical data.

## What it does

- Reads price data from Excel.
- Uses either a `Signal` column or a default 20-day SMA crossover strategy.
- Starts with Rs. 2,00,000 by default.
- Writes an Excel report with summary, trades, equity curve, open positions, and the standardized signal data.

## Expected input

Your Excel file can be either:

- A single sheet with columns like `Date`, `Symbol`, `Open`, `Close`, and optional `Signal`.
- Multiple sheets, where each sheet is one stock and includes `Date`, `Open`, and `Close`.

If `Signal` exists, the script uses it directly:

- `1` or `Buy` means buy.
- `-1` or `Sell` means sell.
- `0` or `Hold` means do nothing.

If `Signal` does not exist, the script creates one using a 20-day simple moving average crossover.

## Run

```powershell
c:/Users/HP/OneDrive/Documents/Desktop/Nepal's stocks/.venv/Scripts/python.exe trade_30_stocks.py data.xlsx --output trading_report.xlsx
```

## Example

```powershell
c:/Users/HP/OneDrive/Documents/Desktop/Nepal's stocks/.venv/Scripts/python.exe trade_30_stocks.py data.xlsx --sheet Prices --initial-cash 200000 --max-positions 30
```

## Notes

- This is an end-of-day style backtest template, not a live trading bot.
- If you want, I can adapt it to your exact Excel layout and strategy rules.