import yfinance as yf
import pandas as pd

def get_return(ticker, months=6):
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period=f"{months+1}mo")  # +1 to zapewnić dane
        if hist.empty or len(hist) < 2:
            return "N/A"

        price_now = hist["Close"].iloc[-1]
        price_then = hist["Close"].iloc[0]

        return_pct = round((price_now - price_then) / price_then * 100, 2)
        return return_pct
    except:
        return "N/A"

def generate_labels(fundamental_csv, output_csv):
    df = pd.read_csv(fundamental_csv)

    labels = []
    for ticker in df["Ticker"]:
        print(f"Pobieranie: {ticker}")
        ret = get_return(ticker, months=6)
        if ret == "N/A":
            labels.append("N/A")
        else:
            labels.append(1 if ret >= 10 else 0)

    df["Target (6m +10%)"] = labels
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    print(f"Zapisano do: {output_csv}")

if __name__ == "__main__":
    generate_labels("../data/analiza_fundamentalna.csv", "../data/z_etykietami.csv")
