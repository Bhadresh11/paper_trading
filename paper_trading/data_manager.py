# data_manager.py
import os
import openpyxl
from openpyxl import Workbook
import requests
from datetime import datetime
import pytz

local_tz = pytz.timezone("Asia/Phnom_Penh")

def make_filename(symbol="ethusdt"):
    ts = datetime.now(local_tz).strftime("%Y%m%d_%H%M%S")
    filename = f"{symbol.upper()}_{ts}.xlsx"
    return filename

def init_workbook(filepath):
    """Create new workbook with required sheets and headers."""
    wb = Workbook()
    # Candles sheet
    s1 = wb.active
    s1.title = "CANDLES"
    s1.append(["Candle Time (UTC)", "Candle Time (Local)", "Open", "High", "Low", "Close", "Volume"])
    # Swings sheet
    s2 = wb.create_sheet("SWINGS")
    s2.append(["Candle Time (UTC)", "Candle Time (Local)", "Swing High", "Swing Low"])
    # Trades sheet
    s3 = wb.create_sheet("TRADES")
    s3.append(["Entry Time (Local)", "Direction", "Entry Price", "Exit Time (Local)", "Exit Price", "Profit", "Stop Loss", "Trail Price"])
    # Stats sheet
    s4 = wb.create_sheet("STATS")
    s4.append(["Run Start (Local)", datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")])
    wb.save(filepath)
    return wb

def load_workbook_if_exists(filepath):
    if os.path.exists(filepath):
        return openpyxl.load_workbook(filepath)
    return None

def save_order_excel(wb, filename, order):
    """
    Save breakout order info to a sheet named "ORDERS".
    Each row: timestamp, local_time, type (UPSIDE/DOWNSIDE), entry_price, qty, SL, TL
    """
    if "ORDERS" not in wb.sheetnames:
        ws = wb.create_sheet("ORDERS")
        ws.append(["UTC Time", "Local Time", "Type", "Entry Price", "Qty", "SL", "TL"])
    else:
        ws = wb["ORDERS"]

    ws.append([
        order["utc_time"],
        order["local_time"],
        order["type"],
        order["entry_price"],
        order["qty"],
        order["stop_loss"],
        order["trail_point"]
    ])
    wb.save(filename)

def fetch_historical_binance(symbol="ETHUSDT", interval="1m", limit=500):
    """One-time historical fetch from Binance REST API."""
    url = f"https://api.binance.com/api/v3/klines?symbol={symbol.upper()}&interval={interval}&limit={limit}"
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        candles = []
        for d in data:
            # d[0] is open time ms, d[1] open, d[2] high, d[3] low, d[4] close, d[5] volume, d[6] close time ms
            utc_dt = datetime.utcfromtimestamp(d[6]/1000.0)
            utc_str = utc_dt.strftime("%Y-%m-%d %H:%M:%S")
            local_dt = pytz.utc.localize(utc_dt).astimezone(local_tz)
            local_str = local_dt.strftime("%Y-%m-%d %H:%M:%S")
            candles.append({
                "utc_time": utc_str,
                "local_time": local_str,
                "open": float(d[1]),
                "high": float(d[2]),
                "low": float(d[3]),
                "close": float(d[4]),
                "volume": float(d[5])
            })
        return candles
    except Exception as e:
        print(f"[DATA MANAGER][ERROR] Historical fetch failed: {e}")
        return []

def save_candle_excel(wb, filepath, candle):
    """Append candle row to CANDLES, avoid duplicate last row."""
    sheet = wb["CANDLES"]
    # check duplicate by comparing last row time
    if sheet.max_row > 1:
        last_time = str(sheet.cell(row=sheet.max_row, column=1).value)
        if last_time == candle["utc_time"]:
            return
    sheet.append([candle["utc_time"], candle["local_time"], candle["open"], candle["high"], candle["low"], candle["close"], candle["volume"]])
    wb.save(filepath)

def save_swing_excel(wb, filepath, candle_time_utc, candle_time_local, swing_high, swing_low):
    sheet = wb["SWINGS"]
    sheet.append([candle_time_utc, candle_time_local, swing_high, swing_low])
    wb.save(filepath)

def save_trade_excel(wb, filepath, trade_record):
    sheet = wb["TRADES"]
    sheet.append([
        trade_record.get("entry_time_local"),
        trade_record.get("direction"),
        trade_record.get("entry_price"),
        trade_record.get("exit_time_local"),
        trade_record.get("exit_price"),
        trade_record.get("profit"),
        trade_record.get("stop_loss"),
        trade_record.get("trail_price")
    ])
    wb.save(filepath)

def update_stats_excel(wb, filepath, stats_dict):
    sheet = wb["STATS"]
    # append summary row
    sheet.append([
        "Total Trades", stats_dict.get("total_trades"),
        "Win Rate %", round(stats_dict.get("win_rate", 0), 2),
        "Avg Profit", round(stats_dict.get("avg_profit", 0), 4),
        "Net P&L", round(stats_dict.get("net_profit", 0), 4)
    ])
    wb.save(filepath)
