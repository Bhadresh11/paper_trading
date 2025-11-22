import csv
from datetime import datetime

def save_price(price: float, filename="prices.csv"):
    with open(filename, mode="a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([datetime.now(), price])

def save_trade(trade: dict, filename="trades.csv"):
    with open(filename, mode="a", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            trade.get("timestamp"),
            trade.get("side"),
            trade.get("entry"),
            trade.get("sl"),
            trade.get("high"),
            trade.get("low")
        ])
