# walidacja_danych.py
import pandas as pd

def validate_data(df):
    alerts = []

    def check_row(row):
        na_count = sum([1 for value in row.values if value == "N/A"])

        if na_count > 5:
            status = "Do sprawdzenia"
        else:
            status = "OK"

        # Błędne dane jeśli kluczowe pola są N/A
        for field in ["Price", "P/E", "ROE (%)", "Debt/Assets"]:
            if field in row and row[field] == "N/A":
                status = "Błędne dane"
                alerts.append({"Ticker": row["Ticker"], "Problem": f"Brak danych: {field}"})
                break

        # Skrajne wartości
        try:
            if status == "OK":
                if float(row["P/E"]) < 0 or float(row["P/E"]) > 100:
                    status = "Do sprawdzenia"
                    alerts.append({"Ticker": row["Ticker"], "Problem": "Podejrzane P/E"})
                if float(row["Debt/Assets"]) > 1.5:
                    status = "Do sprawdzenia"
                    alerts.append({"Ticker": row["Ticker"], "Problem": "Bardzo wysokie Debt/Assets"})
                if float(row["ROE (%)"]) < -100 or float(row["ROE (%)"]) > 100:
                    status = "Do sprawdzenia"
                    alerts.append({"Ticker": row["Ticker"], "Problem": "Podejrzane ROE (%)"})
        except:
            pass

        return status

    df["Status"] = df.apply(check_row, axis=1)

    if alerts:
        pd.DataFrame(alerts).to_csv(r"C:\Portfel2025\alerty_walidacja.csv", index=False, encoding="utf-8-sig")
        print("Zapisano plik alerty_walidacja.csv")

    return df
