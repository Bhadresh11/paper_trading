# trade_engine.py
import uuid

class TradeEngine:
    def __init__(self, qty=1.0, stop_loss_points=20.0, trail_point=1.0):
        """
        qty: trade quantity
        stop_loss_points: initial SL distance from entry
        trail_point: minimum move to shift trailing SL
        """
        self.qty = qty
        self.stop_loss_points = stop_loss_points
        self.trail_point = trail_point
        self.active_trade = None
        self.trade_history = []

    # -----------------------------------------------------------
    # Entry Condition
    # -----------------------------------------------------------
    def try_entry(self, price, last_swing_high=None, last_swing_low=None):
        """
        Logic for breakout entry
        """
        direction = None

        if last_swing_high and price > last_swing_high:
            direction = "buy"

        elif last_swing_low and price < last_swing_low:
            direction = "sell"

        if direction is None:
            return None

        trade_id = str(uuid.uuid4())[:8]
        stop_loss = price - self.stop_loss_points if direction == "buy" else price + self.stop_loss_points

        self.active_trade = {
            "trade_id": trade_id,
            "direction": direction,
            "entry_price": price,
            "stop_loss": stop_loss,
            "trail_price": stop_loss,
            "qty": self.qty
        }

        print(f"⚡ NEW TRADE OPENED ⚡ | Side: {direction} | Entry: {price:.2f} | SL: {stop_loss:.2f}")
        return self.active_trade.copy()

    # -----------------------------------------------------------
    # Trailing Stop + Check SL Hit
    # -----------------------------------------------------------
    def update_active(self, current_price):
        if self.active_trade is None:
            return None

        tr = self.active_trade
        direction = tr["direction"]
        entry = tr["entry_price"]
        sl = tr["stop_loss"]
        tsl = tr["trail_price"]

        tsl_updated = False

        # ---------------- BUY ----------------
        if direction == "buy":

            # Move TSL if price moves up at least trail_point
            if current_price - tsl >= self.trail_point:
                new_tsl = current_price - self.stop_loss_points

                # Do not move TSL above entry+SL zone
                new_tsl = min(new_tsl, entry + self.stop_loss_points)

                tr["trail_price"] = new_tsl
                tr["stop_loss"] = new_tsl
                tsl_updated = True

            # Check SL hit
            if current_price <= sl:
                return self.close_trade(current_price)

        # ---------------- SELL ----------------
        elif direction == "sell":

            if tsl - current_price >= self.trail_point:
                new_tsl = current_price + self.stop_loss_points

                # Do not move TSL below entry-SL zone
                new_tsl = max(new_tsl, entry - self.stop_loss_points)

                tr["trail_price"] = new_tsl
                tr["stop_loss"] = new_tsl
                tsl_updated = True

            if current_price >= sl:
                return self.close_trade(current_price)

        if tsl_updated:
            return {
                "tsl_update": True,
                "trade": tr.copy(),
                "trail": tr["trail_price"],
                "stop_loss": tr["stop_loss"]
            }

        return None

    # -----------------------------------------------------------
    # Close Trade
    # -----------------------------------------------------------
    def close_trade(self, exit_price):
        if self.active_trade is None:
            return None

        tr = self.active_trade
        direction = tr["direction"]
        entry = tr["entry_price"]
        qty = tr["qty"]

        # Profit Calculation
        profit = (exit_price - entry) * qty if direction == "buy" else (entry - exit_price) * qty

        # Balance Calculation
        previous_balance = 100000 if not self.trade_history else self.trade_history[-1]["balance_after"]
        new_balance = previous_balance + profit

        closed_trade = {
            "trade_id": tr["trade_id"],
            "direction": direction,
            "entry_price": entry,
            "exit_price": exit_price,
            "profit": round(profit, 2),
            "balance_after": round(new_balance, 2)
        }

        print(f"📌 TRADE CLOSED 💰 | Side: {direction} | Entry: {entry:.2f} | Exit: {exit_price:.2f} | "
              f"P/L: {profit:.2f} | Balance: {new_balance:.2f}")

        self.trade_history.append(closed_trade)
        self.active_trade = None

        return closed_trade

    # -----------------------------------------------------------
    # Stats For Excel
    # -----------------------------------------------------------
    def get_stats(self):
        wins = [t for t in self.trade_history if t["profit"] > 0]
        losses = [t for t in self.trade_history if t["profit"] <= 0]

        win_rate = (len(wins) / len(self.trade_history) * 100) if self.trade_history else 0
        avg_profit = sum(t["profit"] for t in wins) / len(wins) if wins else 0
        avg_loss = sum(t["profit"] for t in losses) / len(losses) if losses else 0

        return {
            "total_trades": len(self.trade_history),
            "win_rate": round(win_rate, 2),
            "avg_profit": round(avg_profit, 2),
            "avg_loss": round(avg_loss, 2)
        }
