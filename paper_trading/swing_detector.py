from collections import deque

class SwingDetector:
    def __init__(self, lookback=5):
        self.lookback = lookback
        self.prices = deque(maxlen=lookback)

    def add_price(self, price: float):
        self.prices.append(price)
        return self.detect_swing()

    def detect_swing(self):
        if len(self.prices) < self.lookback:
            return None

        middle = self.prices[len(self.prices)//2]
        if middle == max(self.prices):
            return ("high", middle)
        if middle == min(self.prices):
            return ("low", middle)
        return None
