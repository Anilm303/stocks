from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


INITIAL_CASH_DEFAULT = 200_000.0
DEFAULT_SMA_WINDOW = 20
DEFAULT_RSI_BUY_MIN = 45.0
DEFAULT_RSI_BUY_MAX = 70.0
DEFAULT_RSI_SELL_MIN = 35.0
DEFAULT_RSI_SELL_MAX = 75.0


@dataclass
class Position:
    symbol: str
    quantity: int
    entry_price: float
    entry_date: pd.Timestamp


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Backtest a simple trading strategy across up to 30 Nepal stock symbols."
    )
    # Get project root to define default paths
    project_root = Path(__file__).parent.parent
    excel_dir = project_root / "Excel_Files"
    data_dir = project_root / "Data"

    parser.add_argument("input", nargs="?", default=str(data_dir), help="Path to a CSV file, Excel file, or folder with historical prices.")
    parser.add_argument(
        "--output",
        default=str(excel_dir / "trading_report.xlsx"),
        help="Path to the Excel report to create.",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Optional sheet name if the workbook has multiple sheets.",
    )
    parser.add_argument(
        "--initial-cash",
        type=float,
        default=INITIAL_CASH_DEFAULT,
        help="Starting capital in rupees.",
    )
    parser.add_argument(
        "--sma-window",
        type=int,
        default=DEFAULT_SMA_WINDOW,
        help="Simple moving average window used when no Signal column is provided.",
    )
    parser.add_argument(
        "--max-positions",
        type=int,
        default=50,
        help="Maximum number of open positions to hold at once.",
    )
    parser.add_argument(
        "--export-trades",
        action="store_true",
        help="Also export the trades dataframe to trades.csv for analysis.",
    )
    return parser.parse_args()


def load_price_data(path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    if path.is_dir():
        frames = load_from_directory(path)
    elif path.suffix.lower() == ".csv":
        frames = [standardize_frame(pd.read_csv(path), path.stem)]
    else:
        workbook = pd.ExcelFile(path)
        if sheet_name is not None:
            raw_frames = [pd.read_excel(workbook, sheet_name=sheet_name)]
            sheet_names = [sheet_name]
        else:
            sheet_names = workbook.sheet_names
            raw_frames = [pd.read_excel(workbook, sheet_name=name) for name in sheet_names]

        frames = [standardize_frame(frame, current_sheet_name or "Sheet1") for current_sheet_name, frame in zip(sheet_names, raw_frames)]

    data = pd.concat(frames, ignore_index=True)
    data["Date"] = pd.to_datetime(data["Date"], format='mixed')
    data = data.sort_values(["Date", "Symbol"]).reset_index(drop=True)
    return data


def load_from_directory(directory: Path) -> list[pd.DataFrame]:
    frames: list[pd.DataFrame] = []
    for file_path in sorted(directory.iterdir()):
        if file_path.suffix.lower() == ".csv":
            frames.append(standardize_frame(pd.read_csv(file_path), file_path.stem))
        elif file_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            workbook = pd.ExcelFile(file_path)
            sheet_frames = [pd.read_excel(workbook, sheet_name=name) for name in workbook.sheet_names]
            frames.extend(
                standardize_frame(frame, f"{file_path.stem}_{sheet_name or 'Sheet1'}")
                for sheet_name, frame in zip(workbook.sheet_names, sheet_frames)
            )

    if not frames:
        raise ValueError(f"No CSV or Excel files found in {directory}")
    return frames


def standardize_frame(frame: pd.DataFrame, fallback_symbol: str) -> pd.DataFrame:
    columns = {column.lower().strip(): column for column in frame.columns}

    if "date" not in columns:
        raise ValueError("Each sheet must include a Date column.")

    if "symbol" in columns:
        symbol_series = frame[columns["symbol"]].astype(str)
    else:
        symbol_series = pd.Series([fallback_symbol] * len(frame))

    close_column = columns.get("close") or columns.get("ltp") or columns.get("adj close") or columns.get("price")
    if close_column is None:
        raise ValueError("Each sheet must include a Close, Ltp, Adj Close, or Price column.")

    open_column = columns.get("open")
    signal_column = columns.get("signal")

    result = frame.copy()
    result = result.rename(columns={columns["date"]: "Date"})

    if "symbol" in columns:
        result = result.rename(columns={columns["symbol"]: "Symbol"})
    else:
        result["Symbol"] = symbol_series

    if open_column is not None:
        result = result.rename(columns={open_column: "Open"})
    else:
        result["Open"] = pd.NA

    result = result.rename(columns={close_column: "Close"})

    if signal_column is not None:
        result = result.rename(columns={signal_column: "Signal"})

    result["Open"] = pd.to_numeric(result["Open"], errors="coerce")
    result["Close"] = pd.to_numeric(result["Close"], errors="coerce")
    numeric_columns = [
        "High",
        "Low",
        "Volume",
        "Turnover",
        "Daily_Return",
        "Log_Return",
        "SMA_5",
        "SMA_20",
        "EMA_12",
        "EMA_26",
        "RSI_14",
        "MACD",
        "MACD_Signal",
        "ATR_14",
        "BB_Middle",
        "BB_Std",
        "BB_Upper",
        "BB_Lower",
        "OBV",
    ]
    for column in numeric_columns:
        if column in result.columns:
            result[column] = pd.to_numeric(result[column], errors="coerce")

    result = result.dropna(subset=["Date", "Symbol", "Close"]).copy()
    return result


def validate_input_columns(frame: pd.DataFrame) -> None:
    required_columns = {"date", "symbol", "close"}
    present_columns = {column.lower().strip() for column in frame.columns}
    missing_columns = required_columns - present_columns
    if missing_columns:
        raise ValueError(
            "Your sheet should have these columns: Date, Symbol, Close. Optional columns: Open, Signal. Missing: "
            + ", ".join(sorted(missing_columns))
        )


def add_strategy_signals(data: pd.DataFrame, sma_window: int) -> pd.DataFrame:
    if "Signal" in data.columns:
        signal = data["Signal"]
        if signal.dtype.kind in {"i", "u", "f"}:
            data["StrategySignal"] = signal.fillna(0).astype(int).clip(-1, 1)
        else:
            normalized = signal.astype(str).str.strip().str.lower()
            data["StrategySignal"] = normalized.map(
                {
                    "buy": 1,
                    "long": 1,
                    "1": 1,
                    "sell": -1,
                    "exit": -1,
                    "-1": -1,
                    "hold": 0,
                    "0": 0,
                }
            ).fillna(0).astype(int)
        return data

    frames: list[pd.DataFrame] = []
    for symbol, group in data.groupby("Symbol", sort=False):
        ordered = group.sort_values("Date").copy()
        ordered_date_idx = ordered.set_index("Date")

        rolling_min = ordered_date_idx["Close"].rolling("90D").min().values
        rolling_max = ordered_date_idx["Close"].rolling("90D").max().values

        start_date = ordered["Date"].min()
        warmup_period = start_date + pd.Timedelta(days=90)

        # Reverting to Bottom 15% range rule
        range_val = rolling_max - rolling_min
        buy_signal = (ordered["Close"] <= (rolling_min + 0.15 * range_val)) & (ordered["Date"] >= warmup_period)

        ordered["StrategySignal"] = 0
        ordered.loc[buy_signal, "StrategySignal"] = 1
        frames.append(ordered)

    return pd.concat(frames, ignore_index=True)


def backtest(data: pd.DataFrame, initial_cash: float, max_positions: int) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    cash = float(initial_cash)
    positions: dict[str, Position] = {}
    trade_rows: list[dict[str, object]] = []
    equity_rows: list[dict[str, object]] = []

    for date, day in data.groupby("Date", sort=True):
        day = day.sort_values("Symbol")
        day_prices = {row.Symbol: float(row.Close) for row in day.itertuples(index=False)}

        for row in day.itertuples(index=False):
            symbol = str(row.Symbol)
            if symbol not in positions:
                continue
            
            current_price = float(row.Open) if pd.notna(row.Open) else float(row.Close)
            position = positions[symbol]
            
            days_held = (date - position.entry_date).days
            if current_price >= position.entry_price + 20 or days_held > 365:
                positions.pop(symbol)
                cash += position.quantity * current_price
                reason = "SELL: Profit Target Reached (+Rs. 20)" if current_price >= position.entry_price + 20 else "SELL: Time-based Exit (1 Year)"
                trade_rows.append(
                    {
                        "Date": date,
                        "Symbol": symbol,
                        "Side": "SELL",
                        "Quantity": position.quantity,
                        "Price": current_price,
                        "TradeValue": position.quantity * current_price,
                        "CashAfterTrade": cash,
                        "EntryDate": position.entry_date,
                        "EntryPrice": position.entry_price,
                        "Reason": reason
                    }
                )

        # Back to original buy candidate logic (no cooling period)
        buy_candidates = [row for row in day.itertuples(index=False) if int(row.StrategySignal) == 1 and str(row.Symbol) not in positions]
        if buy_candidates and cash > 0:
            import random
            random.seed(int(date.timestamp()))
            random.shuffle(buy_candidates)

            slots_available = max(max_positions - len(positions), 0)
            selected_candidates = buy_candidates[:slots_available]

            current_equity = cash + sum(p.quantity * day_prices.get(s, p.entry_price) for s, p in positions.items())
            per_position_budget = current_equity / max_positions if max_positions > 0 else current_equity
            m = len(selected_candidates)
            if m > 0:
                total_needed = per_position_budget * m
                allocations = [per_position_budget] * m if cash >= total_needed else [cash / m] * m
                for row, alloc in zip(selected_candidates, allocations):
                    symbol = str(row.Symbol)
                    buy_price = float(row.Open) if pd.notna(row.Open) else float(row.Close)
                    if buy_price <= 0 or alloc <= 0 or cash < buy_price:
                        continue

                    quantity = int(min(alloc, cash) // buy_price)
                    quantity = (quantity // 10) * 10
                    if quantity < 10: continue

                    trade_value = quantity * buy_price
                    if trade_value > cash:
                        quantity = int(cash // buy_price)
                        trade_value = quantity * buy_price
                    if quantity <= 0: continue

                    cash -= trade_value
                    positions[symbol] = Position(symbol=symbol, quantity=quantity, entry_price=buy_price, entry_date=pd.Timestamp(date))
                    trade_rows.append({
                        "Date": date, "Symbol": symbol, "Side": "BUY", "Quantity": quantity, "Price": buy_price,
                        "TradeValue": trade_value, "Allocated": alloc, "RemainingAllocation": alloc - trade_value,
                        "CashAfterTrade": cash, "EntryDate": pd.Timestamp(date), "EntryPrice": buy_price,
                    })

        unrealized_value = sum(position.quantity * day_prices.get(symbol, position.entry_price) for symbol, position in positions.items())
        equity_rows.append({"Date": date, "Cash": cash, "PositionsValue": unrealized_value, "Equity": cash + unrealized_value, "OpenPositions": len(positions)})

    # Forced exit logic...
    final_date = data["Date"].max()
    final_prices = data.sort_values("Date").groupby("Symbol", as_index=False).tail(1).set_index("Symbol")["Close"].to_dict()
    for symbol, position in list(positions.items()):
        exit_price = float(final_prices.get(symbol, position.entry_price))
        if exit_price >= position.entry_price:
            cash += position.quantity * exit_price
            trade_rows.append({
                "Date": final_date, "Symbol": symbol, "Side": "FORCED_EXIT", "Quantity": position.quantity, "Price": exit_price,
                "TradeValue": position.quantity * exit_price, "CashAfterTrade": cash, "EntryDate": position.entry_date, "EntryPrice": position.entry_price,
            })
            positions.pop(symbol, None)

    # Formatting and returning DataFrames...
    trade_df = pd.DataFrame(trade_rows)
    equity_df = pd.DataFrame(equity_rows)
    holdings_df = pd.DataFrame([{"Symbol": p.symbol, "Quantity": p.quantity, "EntryDate": p.entry_date, "EntryPrice": p.entry_price, "CurrentPrice": float(final_prices.get(p.symbol, p.entry_price)), "UnrealizedPnL": (float(final_prices.get(p.symbol, p.entry_price)) - p.entry_price) * p.quantity} for p in positions.values()])

    # Date formatting
    for df in [trade_df, equity_df, holdings_df]:
        if not df.empty and "Date" in df.columns: df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m-%d")
        if not df.empty and "EntryDate" in df.columns: df["EntryDate"] = pd.to_datetime(df["EntryDate"]).dt.strftime("%Y-%m-%d")

    return trade_df, equity_df, holdings_df

# Rest of the functions (compute_summary, write_workbook, write_report, main)...
def compute_summary(equity_df: pd.DataFrame, trade_df: pd.DataFrame, initial_cash: float, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    if equity_df.empty: return pd.DataFrame()
    final_equity = float(equity_df["Equity"].iloc[-1])
    total_return = final_equity / initial_cash - 1
    return pd.DataFrame([{"Metric": "Initial Cash", "Value": initial_cash}, {"Metric": "Final Equity", "Value": final_equity}, {"Metric": "Total Return", "Value": total_return}, {"Metric": "Total Trades", "Value": len(trade_df)}])

def write_workbook(output_path: Path, sheet_frames: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in sheet_frames.items():
            frame.to_excel(writer, index=False, sheet_name=sheet_name)
    workbook = load_workbook(output_path)
    financial_columns = {"Price", "BuyPrice", "SellPrice", "TradeValue", "BuyValue", "SellValue", "Allocated", "RemainingAllocation", "CashAfterTrade", "PnL", "UnrealizedPnL", "Cash", "PositionsValue", "Equity", "EntryPrice", "CurrentPrice"}
    number_format = '[$-en-IN]#,##,##,##0'
    for worksheet in workbook.worksheets:
        header_map = {cell.value: cell.column for cell in worksheet[1] if cell.value in financial_columns}
        for column_index in header_map.values():
            column_letter = get_column_letter(column_index)
            for row_index in range(2, worksheet.max_row + 1):
                worksheet[f"{column_letter}{row_index}"].number_format = number_format
    workbook.save(output_path)

def build_buy_trades_df(trade_df: pd.DataFrame) -> pd.DataFrame:
    if trade_df.empty: return trade_df
    buy_df = trade_df[trade_df["Side"] == "BUY"].copy()
    if buy_df.empty: return buy_df
    return buy_df.rename(columns={"Price": "BuyPrice", "TradeValue": "BuyValue"})[["Date", "Symbol", "Quantity", "BuyPrice", "BuyValue", "Allocated", "RemainingAllocation", "CashAfterTrade", "EntryDate"]]

def build_sell_trades_df(trade_df: pd.DataFrame) -> pd.DataFrame:
    if trade_df.empty: return trade_df
    sell_df = trade_df[trade_df["Side"].isin(["SELL", "FORCED_EXIT"])].copy()
    if sell_df.empty: return sell_df
    sell_df = sell_df.rename(columns={"Price": "SellPrice", "TradeValue": "SellValue"})
    sell_df["PnL"] = sell_df["SellValue"] - (sell_df["Quantity"] * sell_df["EntryPrice"])
    return sell_df[["Date", "Symbol", "Quantity", "SellPrice", "SellValue", "PnL", "CashAfterTrade", "EntryDate", "EntryPrice"]]

def write_report(output_path: Path, summary_df: pd.DataFrame, trade_df: pd.DataFrame, equity_df: pd.DataFrame, holdings_df: pd.DataFrame) -> None:
    # Save all components as separate CSV files
    base = output_path.parent
    trade_df.to_csv(base / "trading_report_all_trades.csv", index=False)
    summary_df.to_csv(base / "trading_report_summary.csv", index=False)
    equity_df.to_csv(base / "trading_report_equity_curve.csv", index=False)
    holdings_df.to_csv(base / "trading_report_open_positions.csv", index=False)

    buy_df = build_buy_trades_df(trade_df)
    sell_df = build_sell_trades_df(trade_df)

    buy_df.to_csv(base / "trading_report_buy.csv", index=False)
    sell_df.to_csv(base / "trading_report_sell.csv", index=False)

def main() -> None:
    args = parse_args()
    data = load_price_data(Path(args.input), args.sheet)
    data = add_strategy_signals(data, args.sma_window)
    trade_df, equity_df, holdings_df = backtest(data, args.initial_cash, args.max_positions)
    summary_df = compute_summary(equity_df, trade_df, args.initial_cash, data["Date"].min(), data["Date"].max())
    write_report(Path(args.output), summary_df, trade_df, equity_df, holdings_df)
    print(f"Report written to {args.output}")

if __name__ == "__main__":
    main()
