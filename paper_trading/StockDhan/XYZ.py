# ema_trailing_yf.py
import yfinance as yf
import pandas as pd
from datetime import datetime
import pytz
import os
from openpyxl import load_workbook
import time

# ---------------- USER SETTINGS ----------------
symbol = "HDFCBANK.NS"
interval = "1m"       # allowed: '1m', '2m', '5m', '15m', '30m', '60m', '1d', etc.
history_limit = "1d"  # can be '7d', '1mo', '3mo', etc.

# Strategy parameters
stop_loss_pct = 1         # % stop loss
target_pct = 1.5          # % take profit
risk_per_trade_pct = 1    # % of capital used for P&L scaling
TSL_DISTANCE = 10.0       # trailing stop distance in points (approx)
capital_start = 100000.0

# EMA SETTINGS
EMA_FAST = 9
EMA_SLOW = 21

# ---------------- INTERNALS ----------------
sl_factor = stop_loss_pct / 100.0
tp_factor = target_pct / 100.0
risk_factor = risk_per_trade_pct / 100.0

ts = datetime.now().strftime("%Y%m%d_%H%M%S")
excel_file = f"excel_data/{symbol}/{ts}.xlsx"
tz_bkk = pytz.timezone("Asia/Bangkok")

# In-memory containers
candles = pd.DataFrame()
swings = []     # historical cross events
trades = []     # executed trades

# ---------- Helpers ----------
def round2(v):
    return round(float(v), 2)

def to_bkk(ts):
    return ts.tz_convert(tz_bkk)

def calculate_emas(df):
    df['EMA9'] = df['Close'].ewm(span=EMA_FAST, adjust=False).mean()
    df['EMA21'] = df['Close'].ewm(span=EMA_SLOW, adjust=False).mean()

    return df

def detect_crosses(df):
    crosses = []
    for i in range(1, len(df)):
        # Use iloc + column index to guarantee scalars
        a_ema9 = df.iloc[i-1]['EMA9']
        a_ema21 = df.iloc[i-1]['EMA21']
        b_ema9 = df.iloc[i]['EMA9']
        b_ema21 = df.iloc[i]['EMA21']

        # skip NaNs
        if pd.isna(a_ema9) or pd.isna(a_ema21) or pd.isna(b_ema9) or pd.isna(b_ema21):
            continue

        if (a_ema9 <= a_ema21) and (b_ema9 > b_ema21):
            crosses.append({
                "index": i,
                "type": "GOLDEN",
                "time_bkk": df.index[i],
                "price": df.iloc[i]['Close']
            })
        elif (a_ema9 >= a_ema21) and (b_ema9 < b_ema21):
            crosses.append({
                "index": i,
                "type": "DEAD",
                "time_bkk": df.index[i],
                "price": df.iloc[i]['Close']
            })
    return crosses

def ensure_excel_run_sheets(base_file):
    """Create sheets for this run."""
    sheet_candles = "Candles"
    sheet_swings = "Swings"
    sheet_trades = "Trades"
    if not os.path.exists(base_file):
        os.makedirs(os.path.dirname(base_file), exist_ok=True)
        writer = pd.ExcelWriter(base_file, engine='openpyxl', mode='w')
        pd.DataFrame().to_excel(writer, sheet_name=sheet_candles, index=True)
        pd.DataFrame().to_excel(writer, sheet_name=sheet_swings, index=True)
        pd.DataFrame().to_excel(writer, sheet_name=sheet_trades, index=True)
        writer.close()
    return sheet_candles, sheet_swings, sheet_trades

def write_df_safe(base_file, sheet_name, df):
    """Write DataFrame to Excel sheet safely (tz-naive datetimes + rounded numbers)."""
    df_copy = df.copy()

    # Convert datetime columns to tz-naive
    for col in df_copy.columns:
        if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            df_copy[col] = df_copy[col].dt.tz_localize(None)

    # Also convert index if it's datetime
    if isinstance(df_copy.index, pd.DatetimeIndex):
        df_copy.index = df_copy.index.tz_localize(None)
        df_copy = df_copy.reset_index()  # optional: put the index as column

    # Round numeric columns to 2 decimals
    for c in df_copy.select_dtypes(include='float'):
        df_copy[c] = df_copy[c].round(2)

    # Write to Excel
    with pd.ExcelWriter(base_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_copy.to_excel(writer, sheet_name=sheet_name, index=False)


# ---------- Fetch historical data ----------
print(f"Fetching {symbol} historical data from Yahoo Finance...")
data = yf.download(tickers=symbol, period=history_limit, interval=interval)
# Flatten columns if they are multi-level
if isinstance(data.columns, pd.MultiIndex):
    data.columns = [col[0] for col in data.columns]

if data.empty:
    raise ValueError("No data fetched from Yahoo Finance.")

# Convert index to BKK timezone
# Ensure BKK timezone
if data.index.tz is None:
    data.index = data.index.tz_localize('Asia/Kolkata').tz_convert(tz_bkk)
else:
    data.index = data.index.tz_convert(tz_bkk)

candles = data[['Open','High','Low','Close','Volume']].copy()

candles = calculate_emas(candles)

# ---------- Detect crosses ----------
history_crosses = detect_crosses(candles)
print(f"Detected {len(history_crosses)} historical crosses.")

# ---------- Backtest ----------
def backtest_historical_with_tsl(candles_df, crosses, tsl_distance, capital_start, risk_per_trade_pct):
    capital = capital_start
    hist_trades = []
    for cross in crosses:
        entry_idx = cross['index']
        entry_price = candles_df.iloc[entry_idx]['Close']
        entry_time = candles_df.index[entry_idx]
        side = "LONG" if cross['type'] == "GOLDEN" else "SHORT"

        # Initialize TSL / trackers
        if side == "LONG":
            tsl = entry_price - tsl_distance
            highest = entry_price
            sl = entry_price - tsl_distance
            tp = entry_price * (1 + target_pct/100)
        else:
            tsl = entry_price + tsl_distance
            lowest = entry_price
            sl = entry_price + tsl_distance
            tp = entry_price * (1 - target_pct/100)

        exit_price = None
        exit_time = None
        exit_reason = None

        # iterate forward to find exit
        for j in range(entry_idx + 1, len(candles_df)):
            row = candles_df.iloc[j]
            high = row['High']
            low = row['Low']
            close = row['Close']

            if side == "LONG":
                if high > highest:
                    highest = high
                    tsl = highest - tsl_distance
                if high >= tp:
                    exit_price, exit_time, exit_reason = tp, candles_df.index[j], "TP"
                    break
                if low <= sl:
                    exit_price, exit_time, exit_reason = sl, candles_df.index[j], "SL"
                    break
                if low <= tsl:
                    exit_price, exit_time, exit_reason = tsl, candles_df.index[j], "TSL"
                    break
            else:
                if low < lowest:
                    lowest = low
                    tsl = lowest + tsl_distance
                if low <= tp:
                    exit_price, exit_time, exit_reason = tp, candles_df.index[j], "TP"
                    break
                if high >= sl:
                    exit_price, exit_time, exit_reason = sl, candles_df.index[j], "SL"
                    break
                if high >= tsl:
                    exit_price, exit_time, exit_reason = tsl, candles_df.index[j], "TSL"
                    break

        if exit_price is None:
            exit_price, exit_time, exit_reason = candles_df.iloc[-1]['Close'], candles_df.index[-1], "EndOfHistory"

        if side == "LONG":
            pnl_price = exit_price - entry_price
            pnl_pct = pnl_price / entry_price
        else:
            pnl_price = entry_price - exit_price
            pnl_pct = pnl_price / entry_price

        pnl_amount = capital * risk_per_trade_pct/100 * pnl_pct
        capital_before = capital
        capital_after = capital + pnl_amount
        capital = capital_after

        hist_trades.append({
            "cross_type": cross['type'],
            "entry_time": str(entry_time),
            "entry_price": round(entry_price,2),
            "side": side,
            "SL": round(sl,2),
            "TP": round(tp,2),
            "TSL_at_entry": round(entry_price - tsl_distance if side=="LONG" else entry_price + tsl_distance,2),
            "exit_time": str(exit_time),
            "exit_price": round(exit_price,2),
            "exit_reason": exit_reason,
            "pnl_price": round(pnl_price,2),
            "pnl_pct": round(pnl_pct*100,4),
            "pnl_amount": round(pnl_amount,2),
            "capital_before": round(capital_before,2),
            "capital_after": round(capital_after,2)
        })
    return pd.DataFrame(hist_trades)

# Run backtest
hist_trades_df = backtest_historical_with_tsl(candles, history_crosses, TSL_DISTANCE, capital_start, risk_per_trade_pct)
swings_for_file = hist_trades_df.copy()

# ---------- Excel setup ----------
sheet_candles, sheet_swings, sheet_trades = ensure_excel_run_sheets(excel_file)
write_df_safe(excel_file, sheet_candles, candles)
write_df_safe(excel_file, sheet_swings, swings_for_file)
write_df_safe(excel_file, sheet_trades, pd.DataFrame(trades))

print(f"Backtest complete. Data saved to {excel_file}")


# ---------- Live update / pseudo real-time ----------
FETCH_INTERVAL = 10  # seconds

print("Starting live update every 10 seconds...")

# Keep track of last processed index
last_index = len(candles) - 1

while True:
    try:
        # Fetch latest 1-2 bars to make sure we catch any new candle
        latest_data = yf.download(tickers=symbol, period="1d", interval=interval)
        if isinstance(latest_data.columns, pd.MultiIndex):
            latest_data.columns = [col[0] for col in latest_data.columns]

        # Convert index to BKK tz-naive
        if latest_data.index.tz is None:
            latest_data.index = latest_data.index.tz_localize('Asia/Kolkata').tz_convert(tz_bkk)
        else:
            latest_data.index = latest_data.index.tz_convert(tz_bkk)

        latest_candles = latest_data[['Open','High','Low','Close','Volume']].copy()
        latest_candles = calculate_emas(latest_candles)

        # Append only new rows
        new_rows = latest_candles.iloc[last_index+1:]
        if not new_rows.empty:
            for idx, row in new_rows.iterrows():
                candles.loc[idx] = row
            print(f"\nFetched {len(new_rows)} new candle(s). Last time: {new_rows.index[-1]}")
            last_index = len(candles) - 1

            # Detect crosses on the new candles
            new_crosses = detect_crosses(candles)
            # Keep only crosses that are new (after previous last_index)
            crosses_to_process = [c for c in new_crosses if c['index'] > last_index - len(new_rows)]
            if crosses_to_process:
                for cross in crosses_to_process:
                    print(f"[SWING DETECTED] {cross['type']} at {cross['time_bkk']} price {cross['price']:.2f}")
                    # Call backtest function for this single cross
                    temp_df = backtest_historical_with_tsl(candles, [cross], TSL_DISTANCE, capital_start, risk_per_trade_pct)
                    print(temp_df[['cross_type','entry_time','entry_price','side','exit_time','exit_price','exit_reason','pnl_amount']])
                    
            # Save updated candles to Excel
            write_df_safe(excel_file, sheet_candles, candles)
            # Optional: save swings
            swings_for_file = pd.DataFrame(new_crosses)
            write_df_safe(excel_file, sheet_swings, swings_for_file)

        time.sleep(FETCH_INTERVAL)

    except KeyboardInterrupt:
        print("Stopped by user.")
        break
    except Exception as e:
        print("Error fetching live data:", e)
        time.sleep(FETCH_INTERVAL)
