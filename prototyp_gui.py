# scalona_ocena.py
import pandas as pd
import os
import subprocess

def score_company(row):
    score = 0

    # Fundamentalne
    if row["P/E"] != "N/A" and float(row["P/E"]) < 25:
        score += 2
    if row["ROE (%)"] != "N/A" and float(row["ROE (%)"]) > 15:
        score += 2
    if row["Debt/Assets"] != "N/A" and float(row["Debt/Assets"]) < 0.5:
        score += 1
    if row["PEG"] != "N/A" and 0.5 <= float(row["PEG"]) <= 1.5:
        score += 2
    if row["EV/FCF"] != "N/A" and float(row["EV/FCF"]) < 15:
        score += 2

    # Techniczne
    if row["RSI"] != "N/A":
        rsi = float(row["RSI"])
        if rsi < 30:
            score += 2
        elif rsi > 70:
            score -= 1
    if row["Drop from ATH (%)"] != "N/A" and float(row["Drop from ATH (%)"]) > 10:
        score += 1
    if row["SMA50"] != "N/A" and row["SMA200"] != "N/A":
        if float(row["SMA50"]) > float(row["SMA200"]):
            score += 2

    return score

def classify_company(score):
    if score >= 9:
        return "Dobra i tania"
    elif 6 <= score < 9:
        return "Dobra w dobrej cenie"
    elif 4 <= score < 6:
        return "Spółka średnia"
    else:
        return "Spółka słaba"

def validate_data(df):
    alerts = []

    def check_row(row):
        na_count = sum([1 for value in row.values if value == "N/A"])

        if na_count > 5:
            status = "Do sprawdzenia"
        else:
            status = "OK"

        for field in ["Price", "P/E", "ROE (%)", "Debt/Assets"]:
            if field in row and row[field] == "N/A":
                status = "Błędne dane"
                alerts.append({"Ticker": row["Ticker"], "Problem": f"Brak danych: {field}"})
                break

        try:
            if status == "OK":
                pe = float(row["P/E"])
                da = float(row["Debt/Assets"])
                roe = float(row["ROE (%)"])
                if pe < 0 or pe > 100:
                    status = "Do sprawdzenia"
                    alerts.append({"Ticker": row["Ticker"], "Problem": f"Podejrzane P/E = {pe}"})
                if da > 1.5:
                    status = "Do sprawdzenia"
                    alerts.append({"Ticker": row["Ticker"], "Problem": f"Bardzo wysokie Debt/Assets = {da}"})
                if roe < -100 or roe > 100:
                    status = "Do sprawdzenia"
                    alerts.append({"Ticker": row["Ticker"], "Problem": f"Podejrzane ROE (%) = {roe}"})
        except:
            pass

        return status

    df["Status"] = df.apply(check_row, axis=1)

    return df, alerts

def merge_and_classify(fundamental_file, technical_file, output_file):
    df_fund = pd.read_csv(fundamental_file)
    df_tech = pd.read_csv(technical_file)

    df = pd.merge(df_fund, df_tech, on="Ticker", how="outer")
    df["Score"] = df.apply(score_company, axis=1)
    df["Ocena końcowa"] = df["Score"].apply(classify_company)

    df, alerts = validate_data(df)

    df.sort_values(by="Score", ascending=False, inplace=True)

    cols = list(df.columns)
    for col in ["Score", "Ocena końcowa", "Status"]:
        cols.insert(2, cols.pop(cols.index(col)))
    df = df[cols]

    with pd.ExcelWriter(output_file.replace(".csv", ".xlsx"), engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Ocena", index=False)
        if alerts:
            pd.DataFrame(alerts).to_excel(writer, sheet_name="Alerty", index=False)

        worksheet = writer.sheets["Ocena"]
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len)

    print(f"Zapisano do: {output_file.replace('.csv', '.xlsx')}.")

    try:
        subprocess.run(["start", "excel", output_file.replace(".csv", ".xlsx")], shell=True)
    except:
        print("Nie udało się otworzyć pliku w Excelu.")

if __name__ == "__main__":
    fund_path = r"C:\Portfel2025\analiza_fundamentalna.csv"
    tech_path = r"C:\Portfel2025\analiza_techniczna.csv"
    out_path = r"C:\Portfel2025\scalona_ocena.csv"
    merge_and_classify(fund_path, tech_path, out_path)
