from datetime import datetime
import time
import pandas as pd
import yfinance as yf

SYMBOL = "HDFCBANK.NS"  # NSE stock, yfinance uses .NS suffix
QTY = 100
in_position = None

def fetch_candles():
    """
    Fetch last 30 1-minute candles for the given symbol using yfinance
    """
    try:
        data = yf.download(
            tickers=SYMBOL,
            period="1d",
            interval="1m",
            progress=False
        )
        if data.empty:
            print(f"{datetime.now()} - No data fetched")
            return pd.DataFrame()
        
        data = data.tail(30)  # last 30 minutes
        data = data[['Close']].rename(columns={'Close': 'close'})
        data['close'] = data['close'].astype(float)
        return data
    except Exception as e:
        print(f"{datetime.now()} - Error fetching data: {e}")
        return pd.DataFrame()


def check_ema_signal(df):
    """
    Check for EMA crossover signals
    """
    df['ema_5'] = df['close'].ewm(span=5, adjust=False).mean()
    df['ema_21'] = df['close'].ewm(span=21, adjust=False).mean()

    prev_5, prev_21 = df['ema_5'].iloc[-2], df['ema_21'].iloc[-2]
    curr_5, curr_21 = df['ema_5'].iloc[-1], df['ema_21'].iloc[-1]

    if prev_5 <= prev_21 and curr_5 > curr_21:
        return "BUY"
    if prev_5 >= prev_21 and curr_5 < curr_21:
        return "SELL"
    return None


def place_trade(signal, price):
    """
    Simulated trade execution
    """
    global in_position

    if signal == "BUY" and in_position is None:
        print(f"{datetime.now()} - BUY signal at price {price}")
        # Simulate order execution
        in_position = "LONG"

    elif signal == "SELL" and in_position == "LONG":
        print(f"{datetime.now()} - SELL signal at price {price}")
        # Simulate order execution
        in_position = None


print("EMA Strategy Running")

while True:
    try:
        df = fetch_candles()

        if len(df) < 21:
            print(f"{datetime.now()} - Not enough data. Candles fetched: {len(df)}")
            time.sleep(10)
            continue

        # Latest price
        current_price = df['close'].iloc[-1]

        # Determine signal
        signal = check_ema_signal(df)

        # Print full details
        print("\n" + "="*60)
        print(f"{datetime.now()} - EMA Strategy Running")
        print(f"Current Price: {current_price}")
        print(f"Latest Signal: {signal}")
        print(f"In Position: {in_position}")
        print("\nLast 5 candles:")
        print(df[['close', 'ema_5', 'ema_21']].tail(5))
        print("="*60 + "\n")

        # Place trade if signal exists
        if signal:
            place_trade(signal, current_price)

        time.sleep(10)

    except Exception as e:
        print(f"{datetime.now()} - Error: {e}")
        time.sleep(10)
