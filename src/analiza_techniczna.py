# analiza_techniczna.py
import yfinance as yf
import pandas as pd
import os

def fetch_technical_signals(ticker):
    stock = yf.Ticker(ticker)
    try:
        hist = stock.history(period="6mo")
        if hist.empty:
            return {"Ticker": ticker, "Error": "Brak danych historycznych"}

        close = hist["Close"]
        last = close.iloc[-1]
        ma30 = close.rolling(window=30).mean().iloc[-1]
        max60 = close[-60:].max()
        min60 = close[-60:].min()
        sma50 = close.rolling(window=50).mean().iloc[-1]
        sma100 = close.rolling(window=100).mean().iloc[-1] if len(close) >= 100 else "N/A"
        sma200 = close.rolling(window=200).mean().iloc[-1] if len(close) >= 200 else "N/A"
        ema20 = close.ewm(span=20).mean().iloc[-1]
        ema50 = close.ewm(span=50).mean().iloc[-1]
        ema12 = close.ewm(span=12).mean().iloc[-1]
        ema26 = close.ewm(span=26).mean().iloc[-1]
        vwap = (hist['Volume'] * hist['Close']).cumsum() / hist['Volume'].cumsum()
        vwap_last = vwap.iloc[-1]

        drop_high = round((max60 - last) / max60 * 100, 2)
        swing_high = round((max60 - last) / last * 100, 2)
        swing_low = round((last - min60) / last * 100, 2)
        vwap_diff = round((last - vwap_last) / last * 100, 2)

        # RSI calculation
        delta = close.diff()
        gain = delta.where(delta > 0, 0)
        loss = -delta.where(delta < 0, 0)
        avg_gain = gain.rolling(window=14).mean()
        avg_loss = loss.rolling(window=14).mean()
        rs = avg_gain / avg_loss
        rsi = 100 - (100 / (1 + rs))
        rsi_last = round(rsi.iloc[-1], 2)

        if swing_high < 5:
            zone = "Blisko Swing High"
        elif swing_low < 5:
            zone = "Blisko Swing Low"
        else:
            zone = "Neutral"

        return {
            "Ticker": ticker,
            "Last Price": round(last, 2),
            "30d MA": round(ma30, 2),
            "Max (60d)": round(max60, 2),
            "Drop from ATH (%)": drop_high,
            "SMA50": round(sma50, 2),
            "SMA100": round(sma100, 2) if sma100 != "N/A" else "N/A",
            "SMA200": round(sma200, 2) if sma200 != "N/A" else "N/A",
            "EMA20": round(ema20, 2),
            "EMA50": round(ema50, 2),
            "EMA12": round(ema12, 2),
            "EMA26": round(ema26, 2),
            "VWAP": round(vwap_last, 2),
            "VWAP Diff (%)": vwap_diff,
            "Swing High (%)": swing_high,
            "Swing Low (%)": swing_low,
            "RSI": rsi_last,
            "Strefa": zone,
            "EMA Crossover": "Bullish" if ema12 > ema26 else "Bearish"
        }

    except Exception as e:
        return {"Ticker": ticker, "Error": str(e)}

def analyze_many_from_csv(csv_path, output_path):
    df = pd.read_csv(csv_path)
    tickers = df["Ticker"].drop_duplicates().tolist()

    results = []
    for ticker in tickers:
        print(f"Analiza techniczna: {ticker}")
        result = fetch_technical_signals(ticker)
        results.append(result)

    df_out = pd.DataFrame(results)
    df_out.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"Zapisano do pliku: {output_path}")

if __name__ == "__main__":
    input_csv = os.path.join("C:\\", "Portfel2025", "data", "analiza_fundamentalna.csv")
    output_csv = os.path.join("C:\\", "Portfel2025", "data", "analiza_techniczna.csv")
    analyze_many_from_csv(input_csv, output_csv)
