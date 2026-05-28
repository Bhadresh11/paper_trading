from datetime import datetime
import time
import pandas as pd
import numpy as np
import yfinance as yf

# --- CONFIG ---
SYMBOL = "BAJFINANCE.NS"
EMA_SHORT = 5
EMA_LONG = 21
CHECK_INTERVAL = 10
in_position = None
last_signal = None
last_crossover = None

# ----------------------------
# Fetch today's 1-minute candles
# ----------------------------
def fetch_candles():
    try:
        data = yf.download(
            tickers=SYMBOL,
            period="1d",
            interval="1m",
            progress=False
        )
        if data.empty:
            return pd.DataFrame()

        df = data[['Close']].rename(columns={'Close': 'close'})
        df['close'] = df['close'].astype(float)
        return df.tail(50)  # last 50 minutes
    except Exception as e:
        print(f"{datetime.now()} - Error fetching data: {e}")
        return pd.DataFrame()

# ----------------------------
# EMA Calculation
# ----------------------------
def calculate_emas(df):
    df['ema_short'] = df['close'].ewm(span=EMA_SHORT, adjust=False).mean()
    df['ema_long'] = df['close'].ewm(span=EMA_LONG, adjust=False).mean()
    return df

# ----------------------------
# Helper to get last value safely
# ----------------------------
def get_last_value(series, offset=0):
    val = series.iloc[-1 - offset]
    if isinstance(val, (pd.Series, pd.DataFrame, np.ndarray)):
        val = val.item()
    return float(val)

# ----------------------------
# Detect EMA crossover
# ----------------------------
def detect_crossover(df):
    if len(df) < EMA_LONG:
        return None
    prev_short = get_last_value(df['ema_short'], offset=1)
    prev_long = get_last_value(df['ema_long'], offset=1)
    curr_short = get_last_value(df['ema_short'])
    curr_long = get_last_value(df['ema_long'])

    if prev_short <= prev_long and curr_short > curr_long:
        return "BUY"
    elif prev_short >= prev_long and curr_short < curr_long:
        return "SELL"
    return None

# ----------------------------
# Handle a new crossover
# ----------------------------
def handle_crossover(signal, df):
    global in_position, last_crossover, last_signal

    price = get_last_value(df['close'])
    curr_ema_short = get_last_value(df['ema_short'])
    curr_ema_long = get_last_value(df['ema_long'])
    prev_ema_short = get_last_value(df['ema_short'], offset=1)
    prev_ema_long = get_last_value(df['ema_long'], offset=1)

    last_crossover = {
        "time": datetime.now(),
        "signal": signal,
        "price": price,
        "prev_ema_short": prev_ema_short,
        "prev_ema_long": prev_ema_long,
        "curr_ema_short": curr_ema_short,
        "curr_ema_long": curr_ema_long
    }

    print("\n" + "="*60)
    print(f"{last_crossover['time']} - EMA Crossover Detected!")
    print(f"Symbol: {SYMBOL}")
    print(f"Signal: {signal}")
    print(f"Current Price: {price:.2f}")
    print(f"Previous EMA-{EMA_SHORT}: {prev_ema_short:.2f}, EMA-{EMA_LONG}: {prev_ema_long:.2f}")
    print(f"Current EMA-{EMA_SHORT}: {curr_ema_short:.2f}, EMA-{EMA_LONG}: {curr_ema_long:.2f}")
    print(f"In Position: {in_position}")
    print("="*60 + "\n")

    # Update position
    if signal == "BUY" and in_position is None:
        in_position = "LONG"
    elif signal == "SELL" and in_position == "LONG":
        in_position = None

    last_signal = signal  # mark this signal as last

# ----------------------------
# Main Bot
# ----------------------------
print(f"Starting EMA Crossover Bot for {SYMBOL}")

# Fetch initial data and detect the latest crossover in recent candles
df_today = fetch_candles()
df_today = calculate_emas(df_today)
last_signal = detect_crossover(df_today)
if last_signal:
    handle_crossover(last_signal, df_today)  # print last crossover once

# Live loop
while True:
    try:
        df = fetch_candles()
        if df.empty:
            time.sleep(CHECK_INTERVAL)
            continue

        df = calculate_emas(df)
        current_price = get_last_value(df['close'])
        print(f"{datetime.now()} - Live Price: {current_price:.2f} | In Position: {in_position}")

        signal = detect_crossover(df)
        if signal and signal != last_signal:
            handle_crossover(signal, df)

        time.sleep(CHECK_INTERVAL)

    except KeyboardInterrupt:
        print("Bot stopped by user")
        break
    except Exception as e:
        print(f"{datetime.now()} - Error: {e}")
        time.sleep(CHECK_INTERVAL)
