# analiza_fundamentalna.py
import yfinance as yf
import csv
import os
from datetime import datetime

def fetch_financial_details(ticker):
    stock = yf.Ticker(ticker)
    try:
        fin = stock.financials
        bs = stock.balance_sheet
        info = stock.info

        ebit = fin.loc["EBIT"].iloc[0] if "EBIT" in fin.index else "N/A"
        interest_expense = fin.loc["Interest Expense"].iloc[0] if "Interest Expense" in fin.index else "N/A"
        total_assets = bs.loc["Total Assets"].iloc[0] if "Total Assets" in bs.index else "N/A"
        total_debt = bs.loc["Total Debt"].iloc[0] if "Total Debt" in bs.index else "N/A"

        interest_coverage = round(ebit / abs(interest_expense), 2) if ebit != "N/A" and interest_expense != "N/A" and interest_expense != 0 else "N/A"
        debt_to_assets = round(total_debt / total_assets, 2) if total_debt != "N/A" and total_assets != "N/A" and total_assets != 0 else "N/A"

        eps_ttm = info.get("trailingEps")
        eps_forward = info.get("forwardEps")
        eps_growth = round((eps_ttm - eps_forward) / abs(eps_forward) * 100, 2) if eps_ttm and eps_forward != 0 else "N/A"

        revenue_ttm = info.get("totalRevenue")
        previous_year_revenue = stock.financials.loc["Total Revenue"].iloc[1] if "Total Revenue" in stock.financials.index and stock.financials.shape[1] > 1 else None
        revenue_growth = str(round((revenue_ttm - previous_year_revenue) / abs(previous_year_revenue) * 100, 2)) if revenue_ttm and previous_year_revenue else "N/A"

        return {
            "EBIT": ebit,
            "Interest Expense": interest_expense,
            "Total Assets": total_assets,
            "Total Debt": total_debt,
            "Interest Coverage": interest_coverage,
            "Debt/Assets": debt_to_assets,
            "EPS Growth (%)": eps_growth,
            "Revenue Growth (%)": revenue_growth
        }

    except Exception:
        return {
            "EBIT": "N/A",
            "Interest Expense": "N/A",
            "Total Assets": "N/A",
            "Total Debt": "N/A",
            "Interest Coverage": "N/A",
            "Debt/Assets": "N/A",
            "EPS Growth (%)": "N/A",
            "Revenue Growth (%)": "N/A",
        }

def analyze_multiple_companies(tickers, file_path):
    for ticker in tickers:
        try:
            stock = yf.Ticker(ticker)
            info = stock.info
            financials = fetch_financial_details(ticker)
            enterprise_value = info.get("enterpriseValue")
            free_cashflow = info.get("freeCashflow")

            data = {
                "Date": datetime.today().strftime('%Y-%m-%d'),
                "Company": info.get("longName", "N/A"),
                "Ticker": ticker,
                "Price": info.get("currentPrice", "N/A"),
                "P/E": info.get("trailingPE", "N/A"),
                "PEG": info.get("pegRatio", "N/A"),
                "Price/Sales": info.get("priceToSalesTrailing12Months", "N/A"),
                "Price/Book": info.get("priceToBook", "N/A"),
                "ROE (%)": str(round(info.get("returnOnEquity", 0) * 100, 2)) if info.get("returnOnEquity") else "N/A",
                "ROA (%)": str(round(info.get("returnOnAssets", 0) * 100, 2)) if info.get("returnOnAssets") else "N/A",
                "Operating Margin (%)": str(round(info.get("operatingMargins", 0) * 100, 2)) if info.get("operatingMargins") else "N/A",
                "Gross Margin (%)": str(round(info.get("grossMargins", 0) * 100, 2)) if info.get("grossMargins") else "N/A",
                "Current Ratio": info.get("currentRatio", "N/A"),
                "Quick Ratio": info.get("quickRatio", "N/A"),
                "Beta": info.get("beta", "N/A"),
                "Free Cash Flow": free_cashflow,
                "EV/FCF": str(round(enterprise_value / free_cashflow, 2)) if enterprise_value not in [None, 0] and free_cashflow not in [None, 0] else "N/A",
                "Dividend Yield (%)": str(round(info.get("dividendYield", 0) * 100, 2)) if info.get("dividendYield") and info.get("dividendYield") < 1 else "N/A",
                "Dividend Rate": info.get("dividendRate", "N/A"),
                "EPS Growth (%)": financials.get("EPS Growth (%)", "N/A"),
                "Revenue Growth (%)": financials.get("Revenue Growth (%)", "N/A"),
            }

            data.update(financials)
            file_exists = os.path.isfile(file_path)

            with open(file_path, mode="a", newline="", encoding="utf-8-sig") as file:
                writer = csv.DictWriter(file, fieldnames=data.keys())
                if not file_exists:
                    writer.writeheader()
                writer.writerow(data)

            print(f"Zapisano dane dla: {ticker}")

        except Exception as e:
            print(f"Błąd przy analizie {ticker}: {e}")

if __name__ == "__main__":
    with open("../data/tickers.txt") as f:
        tickers = [line.strip() for line in f if line.strip()]

    output_file = r"../data/analiza_fundamentalna.csv"
    analyze_multiple_companies(tickers, output_file)
