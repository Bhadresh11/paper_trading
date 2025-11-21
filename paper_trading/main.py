import yfinance as yf
import pandas as pd
import time
from datetime import datetime

# ================= CONFIG =================
SYMBOL = "^NSEI"          # Nifty 50 index
INTERVAL = "1m"           # 1-minute candles
SWING_LOOKBACK = 3        # previous candles for swing high
SL_POINTS = 20            # initial stop-loss points
TRAIL_STEP = 1            # trailing stop step
QUANTITY = 1              # virtual quantity

# ================= STATE =================
position = None
candle_highs = []
swing_times = []

# ================= FUNCTIONS =================
def fetch_candles(n=10):
    """Fetch last n 1-minute candles"""
    df = yf.download(SYMBOL, period="1d", interval=INTERVAL, progress=False, auto_adjust=False)
    df = df.dropna()
    return df.tail(n)

def check_breakout(df):
    """Check if last candle breaks previous swing highs"""
    global candle_highs, swing_times
    last_candle = df.iloc[-1]

    # convert all to float safely
    high = float(last_candle["High"]) if not isinstance(last_candle["High"], pd.Series) else float(last_candle["High"].iloc[0])
    low = float(last_candle["Low"]) if not isinstance(last_candle["Low"], pd.Series) else float(last_candle["Low"].iloc[0])
    timestamp = last_candle.name

    candle_highs.append(high)
    swing_times.append(timestamp)
    if len(candle_highs) > SWING_LOOKBACK:
        candle_highs.pop(0)
        swing_times.pop(0)

    if len(candle_highs) == SWING_LOOKBACK:
        prev_highs = [float(h) for h in candle_highs[:-1]]
        if high > max(prev_highs):
            return True, timestamp, high, low

    return False, None, None, None

# ================= MAIN LOOP =================
while True:
    try:
        df = fetch_candles(n=SWING_LOOKBACK + 1)

        last_close = float(df["Close"].iloc[-1]) if not isinstance(df["Close"].iloc[-1], pd.Series) else float(df["Close"].iloc[-1].iloc[0])

        breakout, swing_time, swing_high, swing_low = check_breakout(df)

        if position is None and breakout:
            position = {
                "entry_price": last_close,
                "sl": last_close - SL_POINTS,
                "trail_sl": last_close - SL_POINTS,
                "quantity": QUANTITY
            }
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | ⚡ BUY TRIGGER!")
            print(f"Swing Time: {swing_time} | High: {swing_high:.2f} | Low: {swing_low:.2f}")
            print(f"Entry: {last_close:.2f} | Initial SL: {position['sl']:.2f}")

        if position is not None:
            # update trailing SL
            new_trail_sl = max(position["trail_sl"], last_close - TRAIL_STEP)
            if new_trail_sl != position["trail_sl"]:
                position["trail_sl"] = new_trail_sl
                print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 🔹 Trailing SL moved to {position['trail_sl']:.2f}")

            # check stop loss
            if last_close <= position["trail_sl"]:
                pnl = last_close - position["entry_price"]
                print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | ❌ STOP LOSS HIT at {last_close:.2f} | PnL: {pnl:.2f}")
                position = None
                candle_highs = []
                swing_times = []

        time.sleep(60)

    except Exception as e:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | ERROR: {e}")
        time.sleep(5)
