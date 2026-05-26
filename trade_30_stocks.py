from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


INITIAL_CASH_DEFAULT = 200_000.0
DEFAULT_SMA_WINDOW = 20


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
    parser.add_argument("input", help="Path to the Excel file with historical prices.")
    parser.add_argument(
        "--output",
        default="trading_report.xlsx",
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
        default=30,
        help="Maximum number of open positions to hold at once.",
    )
    return parser.parse_args()


def load_price_data(path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)

    if sheet_name is not None:
        raw_frames = [pd.read_excel(workbook, sheet_name=sheet_name)]
        sheet_names = [sheet_name]
    else:
        sheet_names = workbook.sheet_names
        raw_frames = [pd.read_excel(workbook, sheet_name=name) for name in sheet_names]

    frames: list[pd.DataFrame] = []
    for current_sheet_name, frame in zip(sheet_names, raw_frames):
        standardized = standardize_frame(frame, current_sheet_name or "Sheet1")
        frames.append(standardized)

    data = pd.concat(frames, ignore_index=True)
    data["Date"] = pd.to_datetime(data["Date"])
    data = data.sort_values(["Date", "Symbol"]).reset_index(drop=True)
    return data


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

    result = pd.DataFrame(
        {
            "Date": frame[columns["date"]],
            "Symbol": symbol_series,
            "Open": frame[open_column] if open_column is not None else pd.NA,
            "Close": frame[close_column],
        }
    )

    if signal_column is not None:
        result["Signal"] = frame[signal_column]

    result["Open"] = pd.to_numeric(result["Open"], errors="coerce")
    result["Close"] = pd.to_numeric(result["Close"], errors="coerce")
    result = result.dropna(subset=["Date", "Symbol", "Close"]).copy()
    return result


def validate_input_columns(frame: pd.DataFrame) -> None:
    required_columns = {"date", "symbol", "open", "close", "signal"}
    present_columns = {column.lower().strip() for column in frame.columns}
    missing_columns = required_columns - present_columns
    if missing_columns:
        raise ValueError(
            "Your sheet should have these columns: Date, Symbol, Open, Close, Signal. Missing: "
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
        ordered["SMA"] = ordered["Close"].rolling(window=sma_window, min_periods=sma_window).mean()
        prev_close = ordered["Close"].shift(1)
        prev_sma = ordered["SMA"].shift(1)
        buy_signal = (ordered["Close"] > ordered["SMA"]) & (prev_close <= prev_sma)
        sell_signal = (ordered["Close"] < ordered["SMA"]) & (prev_close >= prev_sma)
        ordered["StrategySignal"] = 0
        ordered.loc[buy_signal, "StrategySignal"] = 1
        ordered.loc[sell_signal, "StrategySignal"] = -1
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
            if int(row.StrategySignal) != -1:
                continue
            symbol = str(row.Symbol)
            if symbol not in positions:
                continue
            exit_price = float(row.Open) if pd.notna(row.Open) else float(row.Close)
            position = positions.pop(symbol)
            cash += position.quantity * exit_price
            trade_rows.append(
                {
                    "Date": date,
                    "Symbol": symbol,
                    "Side": "SELL",
                    "Quantity": position.quantity,
                    "Price": exit_price,
                    "TradeValue": position.quantity * exit_price,
                    "CashAfterTrade": cash,
                    "EntryDate": position.entry_date,
                    "EntryPrice": position.entry_price,
                }
            )

        buy_candidates = [row for row in day.itertuples(index=False) if int(row.StrategySignal) == 1 and str(row.Symbol) not in positions]
        if buy_candidates:
            slots_available = max(max_positions - len(positions), 0)
            selected_candidates = buy_candidates[:slots_available]
            allocation = cash / len(selected_candidates) if selected_candidates else 0.0

            for row in selected_candidates:
                symbol = str(row.Symbol)
                buy_price = float(row.Open) if pd.notna(row.Open) else float(row.Close)
                if buy_price <= 0:
                    continue
                quantity = int(allocation // buy_price)
                if quantity <= 0:
                    continue
                trade_value = quantity * buy_price
                cash -= trade_value
                positions[symbol] = Position(
                    symbol=symbol,
                    quantity=quantity,
                    entry_price=buy_price,
                    entry_date=pd.Timestamp(date),
                )
                trade_rows.append(
                    {
                        "Date": date,
                        "Symbol": symbol,
                        "Side": "BUY",
                        "Quantity": quantity,
                        "Price": buy_price,
                        "TradeValue": trade_value,
                        "CashAfterTrade": cash,
                        "EntryDate": pd.Timestamp(date),
                        "EntryPrice": buy_price,
                    }
                )

        unrealized_value = sum(position.quantity * day_prices.get(symbol, position.entry_price) for symbol, position in positions.items())
        equity_rows.append(
            {
                "Date": date,
                "Cash": cash,
                "PositionsValue": unrealized_value,
                "Equity": cash + unrealized_value,
                "OpenPositions": len(positions),
            }
        )

    final_date = data["Date"].max()
    final_prices = (
        data.sort_values("Date")
        .groupby("Symbol", as_index=False)
        .tail(1)
        .set_index("Symbol")["Close"]
        .to_dict()
    )
    for symbol, position in list(positions.items()):
        exit_price = float(final_prices.get(symbol, position.entry_price))
        cash += position.quantity * exit_price
        trade_rows.append(
            {
                "Date": final_date,
                "Symbol": symbol,
                "Side": "FORCED_EXIT",
                "Quantity": position.quantity,
                "Price": exit_price,
                "TradeValue": position.quantity * exit_price,
                "CashAfterTrade": cash,
                "EntryDate": position.entry_date,
                "EntryPrice": position.entry_price,
            }
        )
        positions.pop(symbol, None)

    trade_columns = ["Date", "Symbol", "Side", "Quantity", "Price", "TradeValue", "CashAfterTrade", "EntryDate", "EntryPrice"]
    trade_df = pd.DataFrame(trade_rows, columns=trade_columns)
    if not trade_df.empty:
        trade_df = trade_df.sort_values(["Date", "Symbol", "Side"]).reset_index(drop=True)
    equity_df = pd.DataFrame(equity_rows)
    if not equity_df.empty:
        equity_df = equity_df.sort_values("Date").reset_index(drop=True)
    holdings_df = pd.DataFrame(
        [
            {
                "Symbol": position.symbol,
                "Quantity": position.quantity,
                "EntryDate": position.entry_date,
                "EntryPrice": position.entry_price,
            }
            for position in positions.values()
        ]
    )

    if not positions:
        equity_df = pd.concat(
            [
                equity_df,
                pd.DataFrame(
                    [
                        {
                            "Date": final_date,
                            "Cash": cash,
                            "PositionsValue": 0.0,
                            "Equity": cash,
                            "OpenPositions": 0,
                        }
                    ]
                ),
            ],
            ignore_index=True,
        )

    return trade_df, equity_df, holdings_df


def compute_summary(equity_df: pd.DataFrame, trade_df: pd.DataFrame, initial_cash: float, start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DataFrame:
    if equity_df.empty:
        final_equity = initial_cash
        max_drawdown = 0.0
    else:
        final_equity = float(equity_df["Equity"].iloc[-1])
        running_max = equity_df["Equity"].cummax()
        drawdown = equity_df["Equity"] / running_max - 1.0
        max_drawdown = float(drawdown.min())

    years = max((end_date - start_date).days / 365.25, 1 / 365.25)
    cagr = (final_equity / initial_cash) ** (1 / years) - 1 if initial_cash > 0 else 0.0
    total_return = final_equity / initial_cash - 1 if initial_cash > 0 else 0.0

    buys = trade_df[trade_df["Side"] == "BUY"] if not trade_df.empty else pd.DataFrame()
    sells = trade_df[trade_df["Side"].isin(["SELL", "FORCED_EXIT"])] if not trade_df.empty else pd.DataFrame()
    realized_trades = min(len(buys), len(sells))
    win_rate = 0.0
    if realized_trades > 0:
        pnl_by_symbol: dict[str, float] = {}
        for _, trade in trade_df.iterrows():
            symbol = str(trade["Symbol"])
            pnl_by_symbol.setdefault(symbol, 0.0)
            if trade["Side"] == "BUY":
                pnl_by_symbol[symbol] -= float(trade["TradeValue"])
            else:
                pnl_by_symbol[symbol] += float(trade["TradeValue"])
        positive = sum(1 for value in pnl_by_symbol.values() if value > 0)
        win_rate = positive / max(len(pnl_by_symbol), 1)

    return pd.DataFrame(
        [
            {"Metric": "Initial Cash", "Value": initial_cash},
            {"Metric": "Final Equity", "Value": final_equity},
            {"Metric": "Total Return", "Value": total_return},
            {"Metric": "CAGR", "Value": cagr},
            {"Metric": "Max Drawdown", "Value": max_drawdown},
            {"Metric": "Total Trades", "Value": len(trade_df)},
            {"Metric": "Win Rate", "Value": win_rate},
        ]
    )


def write_report(output_path: Path, summary_df: pd.DataFrame, trade_df: pd.DataFrame, equity_df: pd.DataFrame, holdings_df: pd.DataFrame, data_df: pd.DataFrame) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        trade_df.to_excel(writer, index=False, sheet_name="Trades")
        equity_df.to_excel(writer, index=False, sheet_name="EquityCurve")
        holdings_df.to_excel(writer, index=False, sheet_name="OpenPositions")
        data_df.to_excel(writer, index=False, sheet_name="Signals")


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    data = load_price_data(input_path, args.sheet)
    validate_input_columns(data.rename(columns={"Date": "date", "Symbol": "symbol", "Open": "open", "Close": "close", "Signal": "signal"}))
    data = add_strategy_signals(data, args.sma_window)
    trade_df, equity_df, holdings_df = backtest(data, args.initial_cash, args.max_positions)

    start_date = data["Date"].min()
    end_date = data["Date"].max()
    summary_df = compute_summary(equity_df, trade_df, args.initial_cash, start_date, end_date)

    write_report(output_path, summary_df, trade_df, equity_df, holdings_df, data)
    print(f"Report written to {output_path.resolve()}")
    print(summary_df.to_string(index=False))


if __name__ == "__main__":
    main()
     