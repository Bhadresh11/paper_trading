# trade_engine.py
from datetime import datetime
import pytz

local_tz = pytz.timezone("Asia/Phnom_Penh")

class TradeEngine:
    def __init__(self, start_balance=100000.0, qty=1, stop_loss_points=20.0, trail_point=1.0):
        self.balance = float(start_balance)
        self.qty = float(qty)
        self.stop_loss_points = float(stop_loss_points)
        self.trail_point = float(trail_point)
        self.active = None  # active trade dict or None
        self.closed_trades = []  # list of trade dicts

   
    def try_entry(self, price, last_swing_high, last_swing_low):
        if self.active is not None:
            return None

        # --- UPSIDE BREAKOUT ---
        if last_swing_high is not None and price > last_swing_high:
            entry_price = price
            stop_loss = entry_price - self.stop_loss_points

            trade = {
                "direction": "LONG",
                "entry_price": entry_price,
                "stop_loss": stop_loss,
                "trail_price": entry_price,
                "qty": self.qty,
                "entry_time_local": datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
            }

            print(
                f"[ENTRY-UP] BREAKOUT ABOVE {last_swing_high} → BUY\n"
                f"  Entry: {entry_price}\n"
                f"  SL:    {stop_loss}\n"
                f"  Qty:   {self.qty}\n"
                f"  Trail: {entry_price}"
            )

            self.active = trade
            return trade

        # --- DOWNSIDE BREAKOUT ---
        if last_swing_low is not None and price < last_swing_low:
            entry_price = price
            stop_loss = entry_price + self.stop_loss_points

            trade = {
                "direction": "SHORT",
                "entry_price": entry_price,
                "stop_loss": stop_loss,
                "trail_price": entry_price,
                "qty": self.qty,
                "entry_time_local": datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
            }

            print(
                f"[ENTRY-DOWN] BREAKOUT BELOW {last_swing_low} → SELL\n"
                f"  Entry: {entry_price}\n"
                f"  SL:    {stop_loss}\n"
                f"  Qty:   {self.qty}\n"
                f"  Trail: {entry_price}"
            )

            self.active = trade
            return trade

        return None

    def update_active(self, price):
        """Update trailing stop; check for SL hit. Returns closed trade dict if closed else None."""
        if self.active is None:
            return None

        direction = self.active["direction"]
        entry = self.active["entry_price"]
        sl = self.active["stop_loss"]
        trail = self.active["trail_price"]

        # move trailing if price moves in favor
        if direction == "LONG":
            if price - trail >= self.trail_point:
                # move SL up by trail_point
                self.active["stop_loss"] += self.trail_point
                self.active["trail_price"] = price
        else:  # SHORT
            if trail - price >= self.trail_point:
                self.active["stop_loss"] -= self.trail_point
                self.active["trail_price"] = price

        # check SL hit
        if direction == "LONG" and price <= self.active["stop_loss"]:
            profit = (price - entry) * self.qty
            self.balance += profit
            self.active.update({
                "exit_price": price,
                "profit": profit,
                "exit_time_local": datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
            })
            closed = self.active
            self.closed_trades.append(closed)
            self.active = None
            return closed

        if direction == "SHORT" and price >= self.active["stop_loss"]:
            profit = (entry - price) * self.qty
            self.balance += profit
            self.active.update({
                "exit_price": price,
                "profit": profit,
                "exit_time_local": datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
            })
            closed = self.active
            self.closed_trades.append(closed)
            self.active = None
            return closed

        return None

    def get_stats(self):
        total = len(self.closed_trades)
        if total == 0:
            return {"total_trades": 0, "win_rate": 0.0, "avg_profit": 0.0, "net_profit": 0.0}
        wins = sum(1 for t in self.closed_trades if t["profit"] > 0)
        net = sum(t["profit"] for t in self.closed_trades)
        avg = net / total
        win_rate = wins / total * 100.0
        return {"total_trades": total, "win_rate": win_rate, "avg_profit": avg, "net_profit": net}
