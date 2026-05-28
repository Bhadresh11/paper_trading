# --- CONFIGURE YOUR CREDENTIALS HERE ---
CLIENT_ID = "1000688527"
ACCESS_TOKEN = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJkaGFuIiwicGFydG5lcklkIjoiIiwiZXhwIjoxNzY4ODk1NDY5LCJpYXQiOjE3Njg4MDkwNjksInRva2VuQ29uc3VtZXJUeXBlIjoiU0VMRiIsIndlYmhvb2tVcmwiOiIiLCJkaGFuQ2xpZW50SWQiOiIxMDAwNjg4NTI3In0.QDq-zDdou5YPNX0Xu9DjgKACDTarr0DssUef_zk1xAfhE8EF8zDWvoAVy6TVoXXk5zIC4Mxt5tMWBF_A9QkGGg"

from datetime import datetime
import time
import pandas as pd
from dhanhq import dhanhq


SYMBOL = "1333"
QTY = 100

dhan = dhanhq(CLIENT_ID, ACCESS_TOKEN)
in_position = None

# Assume SYMBOL, QTY, in_position, and dhan are already defined

def fetch_candles():
    today = datetime.now().strftime("%Y-%m-%d")

    response = dhan.intraday_minute_data(
        security_id=SYMBOL,
        exchange_segment=dhan.NSE,
        instrument_type="EQUITY",
        from_date=today,
        to_date=today
    )
    print(f"Raw API response: {response}")
    candles = response.get("data")
    print(f"Candles fetched: {candles}")

    if not candles:
        return pd.DataFrame()

    if isinstance(candles, list):
        df = pd.DataFrame(candles)
    elif isinstance(candles, dict):
        df = pd.DataFrame([candles])
    else:
        return pd.DataFrame()

    df.columns = [c.lower() for c in df.columns]

    if 'close' not in df.columns:
        return pd.DataFrame()

    df['close'] = df['close'].astype(float)
    return df.tail(30)


def check_ema_signal(df):
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
    global in_position

    if signal == "BUY" and in_position is None:
        print(f"{datetime.now()} - BUY signal at price {price}")
        dhan.place_order(
            security_id=SYMBOL,
            exchange_segment=dhan.NSE,
            transaction_type=dhan.BUY,
            quantity=QTY,
            order_type=dhan.MARKET,
            product_type=dhan.INTRA,
            price=0
        )
        in_position = "LONG"

    elif signal == "SELL" and in_position == "LONG":
        print(f"{datetime.now()} - SELL signal at price {price}")
        dhan.place_order(
            security_id=SYMBOL,
            exchange_segment=dhan.NSE,
            transaction_type=dhan.SELL,
            quantity=QTY,
            order_type=dhan.MARKET,
            product_type=dhan.INTRA,
            price=0
        )
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

        # Calculate EMAs
        df['ema_5'] = df['close'].ewm(span=5, adjust=False).mean()
        df['ema_21'] = df['close'].ewm(span=21, adjust=False).mean()

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
