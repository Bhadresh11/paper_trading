from dhanhq import dhanhq

# --- CONFIGURE YOUR CREDENTIALS HERE ---
client_id = "1000688527"
access_token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzUxMiJ9.eyJpc3MiOiJkaGFuIiwicGFydG5lcklkIjoiIiwiZXhwIjoxNzY4ODk1NDY5LCJpYXQiOjE3Njg4MDkwNjksInRva2VuQ29uc3VtZXJUeXBlIjoiU0VMRiIsIndlYmhvb2tVcmwiOiIiLCJkaGFuQ2xpZW50SWQiOiIxMDAwNjg4NTI3In0.QDq-zDdou5YPNX0Xu9DjgKACDTarr0DssUef_zk1xAfhE8EF8zDWvoAVy6TVoXXk5zIC4Mxt5tMWBF_A9QkGGg"

# Initialize DhanHQ
dhan = dhanhq(client_id, access_token)

# 1. Place an Order (Demo)
# Note: Ensure you have enough funds if not using a sandbox
try:
    response = dhan.place_order(
        security_id='1333', # Example: HDFC Bank
        exchange_segment=dhan.NSE,
        transaction_type=dhan.BUY,
        quantity=100,
        order_type=dhan.MARKET,
        product_type=dhan.INTRA, # Or DELIVERY
        price=0
    )
    print("Order Response:", response)
except Exception as e:
    print("Error:", e)

# 2. Get Positions
positions = dhan.get_positions()
print("Positions:", positions)

# 3. Get Holdings
holdings = dhan.get_holdings()
print("Holdings:", holdings)
