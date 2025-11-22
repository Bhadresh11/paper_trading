# main_bot.py
import asyncio
import json
import websockets
from datetime import datetime, timezone
import pytz
import pandas as pd
import requests
from data_manager import ExcelManager
from trade_engine import TradeEngine

# ------------------------- CONFIG -------------------------
SYMBOL = "ETHUSDT"
INTERVAL = "1m"
LOOKBACK = 15
PING_INTERVAL = 20
PING_TIMEOUT = 20
local_tz = pytz.timezone("Asia/Phnom_Penh")

# ------------------------- Excel Manager -------------------------
excel_mgr = ExcelManager(SYMBOL)

# ------------------------- Memory -------------------------
candles = []

# ------------------------- Historical candles -------------------------
def fetch_historical_binance(symbol, interval, limit=500):
    url = "https://api.binance.com/api/v3/klines"
    params = {"symbol": symbol, "interval": interval, "limit": limit}
    r = requests.get(url, params=params)
    if r.status_code == 200:
        data = r.json()
        candle_list = []
        for k in data:
            utc_dt = datetime.fromtimestamp(k[6]/1000.0, tz=timezone.utc)
            local_dt = utc_dt.astimezone(local_tz)
            candle = {
                "utc_time": utc_dt.strftime("%Y-%m-%d %H:%M:%S"),
                "local_time": local_dt.strftime("%Y-%m-%d %H:%M:%S"),
                "open": float(k[1]),
                "high": float(k[2]),
                "low": float(k[3]),
                "close": float(k[4]),
                "volume": float(k[5])
            }
            candle_list.append(candle)
        return candle_list
    return []

# Load historical
print("[MAIN] Loading historical candles...")
historical = fetch_historical_binance(SYMBOL, INTERVAL, limit=500)
if historical:
    candles.extend(historical)
    for c in historical:
        excel_mgr.save_candle(c)
    print(f"[MAIN] {len(historical)} historical candles loaded and saved.")
else:
    print("[MAIN] No historical candles found.")

# ------------------------- Swing calculation -------------------------
def find_last_swings(df):
    # Convert string to datetime
    df['utc_time'] = pd.to_datetime(df['utc_time'])
    df['local_time'] = pd.to_datetime(df['local_time'])

    # Swing high / low
    swing_high_value = df['high'].max()
    swing_low_value = df['low'].min()

    # Times of swing high / low
    swing_high_row = df.loc[df['high'].idxmax()]
    swing_low_row = df.loc[df['low'].idxmin()]

    swing_high_time_utc = swing_high_row['utc_time']
    swing_high_time_local = swing_high_row['local_time']
    swing_low_time_utc = swing_low_row['utc_time']
    swing_low_time_local = swing_low_row['local_time']

    return swing_high_value, swing_high_time_utc, swing_high_time_local, swing_low_value, swing_low_time_utc, swing_low_time_local


# Convert historical to DataFrame
historical_df = pd.DataFrame(historical)
swing_high, swing_high_utc, swing_high_local, swing_low, swing_low_utc, swing_low_local = find_last_swings(historical_df)

print("[MAIN] Last Swing Levels:")
print(f"  ➤ Swing High = {swing_high} | UTC: {swing_high_utc} | Local: {swing_high_local}")
print(f"  ➤ Swing Low  = {swing_low} | UTC: {swing_low_utc} | Local: {swing_low_local}")

# Save initial swings to Excel
excel_mgr.save_swing(
    str(swing_high_utc), str(swing_high_local), swing_high,
    str(swing_low_utc), str(swing_low_local), swing_low
)

# ------------------------- Trade Engine -------------------------
engine = TradeEngine(qty=1.0, stop_loss_points=20.0, trail_point=1.0)
print("[MAIN] TradeEngine initialized.")

last_executed_up_price = None
last_executed_down_price = None

# ------------------------- WebSocket loop -------------------------
async def ws_loop():
    global swing_high, swing_low, last_executed_up_price, last_executed_down_price

    ws_url = f"wss://stream.binance.com:9443/ws/{SYMBOL.lower()}@kline_{INTERVAL}"
    while True:
        try:
            print("[WS] Connecting...")
            async with websockets.connect(ws_url, ping_interval=PING_INTERVAL, ping_timeout=PING_TIMEOUT) as ws:
                print("[WS] Connected OK.")
                async for message in ws:
                    data = json.loads(message)
                    k = data.get("k", {})
                    if not k.get("x", False):
                        continue

                    # timestamps
                    utc_dt = datetime.fromtimestamp(k["T"]/1000.0, tz=timezone.utc)
                    local_dt = utc_dt.astimezone(local_tz)
                    utc_str = utc_dt.strftime("%Y-%m-%d %H:%M:%S")
                    local_str = local_dt.strftime("%Y-%m-%d %H:%M:%S")

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
                        excel_mgr.save_candle(new_candle)
                        print(f"[DATA] {local_str} Close={new_candle['close']} saved.")

                    price = new_candle["close"]

                    # Print swing pending levels
                    print(f"[ORDERS PENDING] {local_str} UPSIDE > {swing_high} | DOWNSIDE < {swing_low}")

                    # ------------------------- Check breakout -------------------------
                    if swing_high and price > swing_high:
                        if last_executed_up_price != swing_high and engine.active_trade is None:
                            entry = engine.try_entry(price=price, last_swing_high=swing_high, last_swing_low=swing_low)
                            if entry:
                                excel_mgr.save_order(entry)
                                last_executed_up_price = swing_high
                                # recalc swing high
                                swing_high = max([c["high"] for c in candles[-LOOKBACK:]])
                                excel_mgr.save_swing(utc_str, local_str, swing_high, swing_low)

                    if swing_low and price < swing_low:
                        if last_executed_down_price != swing_low and engine.active_trade is None:
                            entry = engine.try_entry(price=price, last_swing_high=swing_high, last_swing_low=swing_low)
                            if entry:
                                excel_mgr.save_order(entry)
                                last_executed_down_price = swing_low
                                # recalc swing low
                                swing_low = min([c["low"] for c in candles[-LOOKBACK:]])
                                excel_mgr.save_swing(utc_str, local_str, swing_high, swing_low)

                    # ------------------------- Update active trade -------------------------
                    result = engine.update_active(price)
                    if result is None:
                        if engine.active_trade:
                            tr = engine.active_trade
                            print(f"[TSL] {local_str} Active {tr['direction']} TSL={tr['trail_price']} SL={tr['stop_loss']} Entry={tr['entry_price']} Current={price}")
                            excel_mgr.save_tsl_update({
                                "time_local": local_str,
                                "trade_id": tr["trade_id"],
                                "direction": tr["direction"],
                                "current_price": price,
                                "trail_price": tr["trail_price"],
                                "stop_loss": tr["stop_loss"]
                            })
                    else:
                        if isinstance(result, dict) and result.get("tsl_update"):
                            tr = result["trade"]
                            print(f"[TSL MOVE] {local_str} {tr['direction']} → New TSL = {result['trail']}, New SL = {result['stop_loss']}")
                            excel_mgr.save_tsl_update({
                                "time_local": local_str,
                                "trade_id": tr["trade_id"],
                                "direction": tr["direction"],
                                "current_price": price,
                                "trail_price": result["trail"],
                                "stop_loss": result["stop_loss"]
                            })
                        else:
                            closed = result
                            profit = closed.get("profit", 0)
                            emoji = "✅" if profit > 0 else "❌"
                            sign = "+" if profit > 0 else ""
                            closed["note"] = f"{emoji} {sign}{round(profit,4)}"
                            print(f"[TRADE EXIT] {local_str} {closed['direction']} exit @ {closed['exit_price']:.4f} Profit={sign}{closed['profit']:.4f} {emoji} Balance={closed['balance_after']:.4f}")
                            excel_mgr.save_order(closed)
                            stats = engine.get_stats()
                            excel_mgr.save_stats(stats)

        except Exception as e:
            import traceback
            print(f"[MAIN][WS ERROR] {e}\n{traceback.format_exc()}")
            print("[MAIN] Reconnecting in 5s...")
            await asyncio.sleep(5)

# ------------------------- Run -------------------------
if __name__ == "__main__":
    asyncio.run(ws_loop())
