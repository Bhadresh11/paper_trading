import requests
import time

API_KEY = "d5mv84hr01qj2afib7igd5mv84hr01qj2afib7j0"
SYMBOL = "NSE:RELIANCE"

while True:
    try:
        url = f"https://finnhub.io/api/v1/quote?symbol={SYMBOL}&token={API_KEY}"
        data = requests.get(url).json()
        current_price = data['c']  # 'c' = current price
        print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - Current Price: {current_price}")
        time.sleep(1)
    except Exception as e:
        print("Error:", e)
        time.sleep(5)
