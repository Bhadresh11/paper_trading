import websocket
import json
import pandas as pd
from datetime import datetime
import pytz
import os

# Parameters
symbol = "BTCUSDT"
interval = "1m"
stop_loss_pct = 0.02   # 2% stop-loss
target_pct = 0.035     # 3.5% target
capital = 100000       # Starting capital
risk_per_trade_pct = 0.01

# Bangkok timezone
tz_bkk = pytz.timezone('Asia/Bangkok')

# Data storage
candles = pd.DataFrame(columns=['open','high','low','close','volume','EMA9','EMA21','candle_time_bkk'])
trades = []

# Excel file path
excel_file = "EMA_9_21_BTC_Algo.xlsx"

def calculate_ema(df, period):
    return df['close'].ewm(span=period, adjust=False).mean()

def log_trade(trade):
    print(f"[{trade['candle_time_bkk']}] {trade['type']} ENTRY: {trade['entry']}, SL: {trade['SL']}, TP: {trade['TP']}, Capital: {trade['capital']}")

def save_to_excel():
    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
        candles.to_excel(writer, sheet_name='Candle_Data', index=False)
        if trades:
            trades_df = pd.DataFrame(trades)
            trades_df.to_excel(writer, sheet_name='Trade_Log', index=False)

def on_message(ws, message):
    global candles, capital
    msg = json.loads(message)
    kline = msg['k']
    
    if kline['x']:  # candle closed
        candle_time = datetime.fromtimestamp(kline['t']/1000, pytz.utc).astimezone(tz_bkk)
        new_row = {
            'open': float(kline['o']),
            'high': float(kline['h']),
            'low': float(kline['l']),
            'close': float(kline['c']),
            'volume': float(kline['v']),
            'candle_time_bkk': candle_time,
            'EMA9': None,
            'EMA21': None
        }
        candles.loc[len(candles)] = new_row
        print(f"Close Price: {new_row}")

        # Calculate EMAs
        if len(candles) > 21:
            candles['EMA9'] = calculate_ema(candles, 9)
            candles['EMA21'] = calculate_ema(candles, 21)
            
            last = candles.iloc[-1]
            prev = candles.iloc[-2]
            
            # Print candle close price
            print(f"[{last['candle_time_bkk']}] Close Price: {last['close']:.2f}, EMA9: {last['EMA9']:.2f}, EMA21: {last['EMA21']:.2f}")
            
            # Detect EMA crossover
            if prev['EMA9'] < prev['EMA21'] and last['EMA9'] > last['EMA21']:
                entry_price = last['close']
                sl = entry_price * (1 - stop_loss_pct)
                tp = entry_price * (1 + target_pct)
                trade = {
                    'candle_time_bkk': last['candle_time_bkk'],
                    'type': 'LONG',
                    'entry': entry_price,
                    'SL': sl,
                    'TP': tp,
                    'status': 'OPEN',
                    'capital': capital
                }
                trades.append(trade)
                log_trade(trade)
            
            elif prev['EMA9'] > prev['EMA21'] and last['EMA9'] < last['EMA21']:
                entry_price = last['close']
                sl = entry_price * (1 + stop_loss_pct)
                tp = entry_price * (1 - target_pct)
                trade = {
                    'candle_time_bkk': last['candle_time_bkk'],
                    'type': 'SHORT',
                    'entry': entry_price,
                    'SL': sl,
                    'TP': tp,
                    'status': 'OPEN',
                    'capital': capital
                }
                trades.append(trade)
                log_trade(trade)
        
        # Update trades for TP/SL hit
        for t in trades:
            if t['status'] == 'OPEN':
                if t['type'] == 'LONG':
                    if last['high'] >= t['TP']:
                        t['status'] = 'TP HIT'
                        capital += capital * risk_per_trade_pct * target_pct
                        print(f"[{last['candle_time_bkk']}] LONG TP HIT at {t['TP']:.2f}, New Capital: {capital:.2f}")
                    elif last['low'] <= t['SL']:
                        t['status'] = 'SL HIT'
                        capital -= capital * risk_per_trade_pct * stop_loss_pct
                        print(f"[{last['candle_time_bkk']}] LONG SL HIT at {t['SL']:.2f}, New Capital: {capital:.2f}")
                elif t['type'] == 'SHORT':
                    if last['low'] <= t['TP']:
                        t['status'] = 'TP HIT'
                        capital += capital * risk_per_trade_pct * target_pct
                        print(f"[{last['candle_time_bkk']}] SHORT TP HIT at {t['TP']:.2f}, New Capital: {capital:.2f}")
                    elif last['high'] >= t['SL']:
                        t['status'] = 'SL HIT'
                        capital -= capital * risk_per_trade_pct * stop_loss_pct
                        print(f"[{last['candle_time_bkk']}] SHORT SL HIT at {t['SL']:.2f}, New Capital: {capital:.2f}")
        
        # Save to Excel after every candle
        save_to_excel()

# Binance WebSocket URL
ws_url = f"wss://stream.binance.com:9443/ws/{symbol.lower()}@kline_{interval}"

print(f"Starting EMA 9–21 Algo for {symbol} with Capital: {capital} THB")
ws = websocket.WebSocketApp(ws_url, on_message=on_message)
ws.run_forever()
