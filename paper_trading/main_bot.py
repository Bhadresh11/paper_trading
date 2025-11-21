# main_bot.py
import asyncio
import json
import websockets
from datetime import datetime, timezone
import pytz
from data_manager import make_filename, init_workbook, fetch_historical_binance, save_candle_excel, save_swing_excel, save_trade_excel, update_stats_excel
from trade_engine import TradeEngine

# CONFIG
SYMBOL = "ETHUSDT"
INTERVAL = "1m"
LOOKBACK = 15
PING_INTERVAL = 20
PING_TIMEOUT = 20

local_tz = pytz.timezone("Asia/Phnom_Penh")

# Excel file
FILENAME = make_filename(SYMBOL)
wb = init_workbook(FILENAME)

# Memory storage
candles = []

# Historical data
print("[MAIN] Loading historical data ONCE...")
historical = fetch_historical_binance(SYMBOL, INTERVAL, limit=500)
if historical:
    candles.extend(historical)
    print(f"[MAIN] {len(historical)} historical candles loaded.")
    for c in historical:
        save_candle_excel(wb, FILENAME, c)
else:
    print("[MAIN] No historical candles; continuing live only.")

# Swing initialization
last_swing_high, last_swing_low = None, None
if len(candles) >= LOOKBACK:
    highs = [c["high"] for c in candles[-LOOKBACK:]]
    lows = [c["low"] for c in candles[-LOOKBACK:]]
    last_swing_high = max(highs)
    last_swing_low = min(lows)
    print(f"[MAIN] Initial Last Swing High={last_swing_high}, Last Swing Low={last_swing_low}")
    save_swing_excel(wb, FILENAME, candles[-1]["utc_time"], candles[-1]["local_time"], last_swing_high, last_swing_low)

# Trade engine
engine = TradeEngine(start_balance=100000.0, qty=1.0, stop_loss_points=20.0, trail_point=1.0)
print("[MAIN] TradeEngine initialized.")

def detect_swing_direction(prev_high, prev_low, new_high, new_low):
    if prev_high is None or new_high > prev_high:
        return "UPSIDE"
    if prev_low is None or new_low < prev_low:
        return "DOWNSIDE"
    return "NEUTRAL"

async def ws_loop():
    global last_swing_high, last_swing_low
    ws_url = f"wss://stream.binance.com:9443/ws/{SYMBOL.lower()}@kline_{INTERVAL}"
    while True:
        try:
            print("[WS] Connecting...")
            async with websockets.connect(ws_url, ping_interval=PING_INTERVAL, ping_timeout=PING_TIMEOUT, close_timeout=5, max_size=None) as ws:
                print("[WS] Connected OK.")
                async for message in ws:
                    data = json.loads(message)
                    k = data.get("k", {})
                    if not k.get("x", False):
                        continue

                    # Datetime
                    utc_dt = datetime.fromtimestamp(k["T"]/1000.0, tz=timezone.utc)
                    local_dt = utc_dt.astimezone(local_tz)
                    utc_str = utc_dt.strftime("%Y-%m-%d %H:%M:%S")
                    local_str = local_dt.strftime("%Y-%m-%d %H:%M:%S")

                    # Candle
                    new_candle = {
                        "utc_time": utc_str,
                        "local_time": local_str,
                        "open": float(k["o"]),
                        "high": float(k["h"]),
                        "low": float(k["l"]),
                        "close": float(k["c"]),
                        "volume": float(k.get("v", 0))
                    }

                    if len(candles) == 0 or candles[-1]["utc_time"] != new_candle["utc_time"]:
                        candles.append(new_candle)
                        save_candle_excel(wb, FILENAME, new_candle)
                        print(f"[DATA] {local_str} Close={new_candle['close']} saved.")

                    # Compute swing
                    if len(candles) >= LOOKBACK:
                        recent = candles[-LOOKBACK:]
                        swing_high = max(c["high"] for c in recent)
                        swing_low = min(c["low"] for c in recent)
                        direction = detect_swing_direction(last_swing_high, last_swing_low, swing_high, swing_low)

                        # New swing
                        if last_swing_high is None or swing_high != last_swing_high or swing_low != last_swing_low:
                            print(f"[SWING] New swing detected at {local_str}: High={swing_high}, Low={swing_low} | Direction: {direction}")
                            save_swing_excel(wb, FILENAME, new_candle["utc_time"], new_candle["local_time"], swing_high, swing_low)
                            last_swing_high, last_swing_low = swing_high, swing_low

                        # Pending breakout orders
                        print(f"[ORDERS PENDING] {local_str} UPSIDE breakout > {last_swing_high:.2f}")
                        print(f"[ORDERS PENDING] {local_str} DOWNSIDE breakout < {last_swing_low:.2f}")

                        price = new_candle["close"]

                        # Check actual breakout for UPSIDE
                        if price > last_swing_high:
                            entry = engine.try_entry(price=price, last_swing_high=last_swing_high, last_swing_low=last_swing_low)
                            if entry:
                                print(f"[ORDERS] {local_str} UPSIDE breakout > {last_swing_high:.2f} → BUY Qty={entry['qty']} SL={entry['stop_loss']:.2f} TL={entry['trail_price']:.2f}")
                                save_trade_excel(wb, FILENAME, entry)

                        # Check actual breakout for DOWNSIDE
                        if price < last_swing_low:
                            entry = engine.try_entry(price=price, last_swing_high=last_swing_high, last_swing_low=last_swing_low)
                            if entry:
                                print(f"[ORDERS] {local_str} DOWNSIDE breakout < {last_swing_low:.2f} → SELL Qty={entry['qty']} SL={entry['stop_loss']:.2f} TL={entry['trail_price']:.2f}")
                                save_trade_excel(wb, FILENAME, entry)

                        # Update active trade (trailing stop)
                        closed = engine.update_active(price=price)
                        if closed:
                            print(f"[TRADE EXIT] {closed['exit_time_local']} {closed['direction']} exit @ {closed['exit_price']:.4f} Profit={closed['profit']:.4f} Balance={engine.balance:.4f}")
                            save_trade_excel(wb, FILENAME, closed)
                            update_stats_excel(wb, FILENAME, engine.get_stats())

        except Exception as e:
            import traceback
            print(f"[MAIN][WS ERROR] {e}\n{traceback.format_exc()}")
            print("[MAIN] Reconnecting in 5s...")
            await asyncio.sleep(5)

if __name__ == "__main__":
    asyncio.run(ws_loop())
