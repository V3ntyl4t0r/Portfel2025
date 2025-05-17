# advanced_tech_analysis.py
import yfinance as yf
import pandas as pd

def analyze_advanced_signals(ticker):
    stock = yf.Ticker(ticker)
    try:
        hist = stock.history(period="3mo")  # 3 miesiące
        if hist.empty:
            return {
                "Swing High (%)": "N/A",
                "Swing Low (%)": "N/A",
                "EMA12/26": "N/A",
                "VWAP Diff (%)": "N/A"
            }

        close = hist["Close"]
        high = hist["High"]
        low = hist["Low"]

        current_price = close.iloc[-1]

        # Swing High/Low
        swing_high = high[-60:].max()
        swing_low = low[-60:].min()
        swing_high_diff = round((current_price - swing_high) / swing_high * 100, 2)
        swing_low_diff = round((current_price - swing_low) / swing_low * 100, 2)

        # EMA12 / EMA26
        ema12 = close.ewm(span=12, adjust=False).mean().iloc[-1]
        ema26 = close.ewm(span=26, adjust=False).mean().iloc[-1]
        ema_status = "Wzrost (EMA12 > EMA26)" if ema12 > ema26 else "Spadek (EMA12 < EMA26)"

        # VWAP HLC3
        hlc3 = (high + low + close) / 3
        vwap = (hlc3 * hist["Volume"]).cumsum() / hist["Volume"].cumsum()
        vwap_diff = round((current_price - vwap.iloc[-1]) / vwap.iloc[-1] * 100, 2)

        return {
            "Swing High (%)": swing_high_diff,
            "Swing Low (%)": swing_low_diff,
            "EMA12/26": ema_status,
            "VWAP Diff (%)": vwap_diff
        }

    except Exception as e:
        return {
            "Swing High (%)": "N/A",
            "Swing Low (%)": "N/A",
            "EMA12/26": "N/A",
            "VWAP Diff (%)": "N/A"
        }
