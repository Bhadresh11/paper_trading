import yfinance as yf
import pandas as pd
import numpy as np

# 1. Download 1-minute ETHUSD data
def load_data(symbol="ETH-USD", interval="1m", period="1d"):
    df = yf.download(symbol, interval=interval, period=period)
    df.dropna(inplace=True)
    return df

# 2. Detect swing highs/lows using numpy arrays
def detect_swings(df, left=2, right=2):
    highs = df["High"].values
    lows = df["Low"].values
    n = len(df)

    swing_high = np.zeros(n, dtype=bool)
    swing_low = np.zeros(n, dtype=bool)

    for i in range(left, n - right):
        is_high = True
        is_low = True

        for j in range(1, left + 1):
            if highs[i] <= highs[i - j]:
                is_high = False
            if lows[i] >= lows[i - j]:
                is_low = False
        for j in range(1, right + 1):
            if highs[i] <= highs[i + j]:
                is_high = False
            if lows[i] >= lows[i + j]:
                is_low = False

        swing_high[i] = is_high
        swing_low[i] = is_low

    df["SwingHigh"] = swing_high
    df["SwingLow"] = swing_low
    return df

# 3. Detect breakout and sweep signals using numpy arrays
def breakout_signals(df):
    highs = df["High"].values
    lows = df["Low"].values
    closes = df["Close"].values
    swing_highs = df["SwingHigh"].values
    swing_lows = df["SwingLow"].values

    n = len(df)
    last_high = None
    last_low = None

    breakout_up = np.zeros(n, dtype=bool)
    breakout_down = np.zeros(n, dtype=bool)
    sweep_up = np.zeros(n, dtype=bool)
    sweep_down = np.zeros(n, dtype=bool)

    for i in range(n):
        # update last swings
        if swing_highs[i]:
            last_high = highs[i]
        if swing_lows[i]:
            last_low = lows[i]

        # bullish breakout
        if last_high is not None:
            if closes[i] > last_high:
                breakout_up[i] = True
            if highs[i] > last_high and closes[i] < last_high:
                sweep_up[i] = True

        # bearish breakout
        if last_low is not None:
            if closes[i] < last_low:
                breakout_down[i] = True
            if lows[i] < last_low and closes[i] > last_low:
                sweep_down[i] = True

    df["BreakoutUp"] = breakout_up
    df["BreakoutDown"] = breakout_down
    df["SweepHigh"] = sweep_up
    df["SweepLow"] = sweep_down

    return df

# 4. Run everything
df = load_data("ETH-USD", interval="1m", period="1d")
df = detect_swings(df)
df = breakout_signals(df)

print(df.tail(50)[[
    "High","Low","SwingHigh","SwingLow","BreakoutUp","BreakoutDown","SweepHigh","SweepLow"
]])
