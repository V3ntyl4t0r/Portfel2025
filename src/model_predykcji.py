import pandas as pd
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import OneHotEncoder
from sklearn.impute import SimpleImputer
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer

def load_and_prepare_data(csv_path):
    df = pd.read_excel(csv_path)

    # Filtrujemy tylko przypadki z targetem
    df = df[df["Target (6m +10%)"].isin([0, 1])]

    # Wybór cech do modelu
    features = [
        "P/E", "PEG", "ROE (%)", "Debt/Assets", "EV/FCF",
        "EPS Growth (%)", "Revenue Growth (%)", "RSI", "Drop from ATH (%)",
        "SMA50", "SMA200", "Beta", "Dividend Yield (%)",
        "EMA Crossover", "Strefa"
    ]

    X = df[features]
    y = df["Target (6m +10%)"]

    # Podział na cechy numeryczne i kategoryczne
    numeric_features = [col for col in features if col not in ["EMA Crossover", "Strefa"]]
    categorical_features = ["EMA Crossover", "Strefa"]

    # Pipeline dla danych numerycznych i kategorii
    numeric_pipeline = Pipeline([
        ("imputer", SimpleImputer(strategy="mean"))
    ])

    categorical_pipeline = Pipeline([
        ("encoder", OneHotEncoder(handle_unknown="ignore"))
    ])

    preprocessor = ColumnTransformer([
        ("num", numeric_pipeline, numeric_features),
        ("cat", categorical_pipeline, categorical_features)
    ])

    return df, preprocessor, X, y

def train_model(X, y, preprocessor):
    # Tworzenie końcowego pipeline
    model = Pipeline([
        ("preprocessor", preprocessor),
        ("classifier", RandomForestClassifier(n_estimators=100, random_state=42))
    ])

    model.fit(X, y)
    return model

def run_prediction(csv_path, output_path):
    df, preprocessor, X, y = load_and_prepare_data(csv_path)
    model = train_model(X, y, preprocessor)

    # Predykcja
    df["ML_Predicted"] = model.predict(X)

    # Eksport wyników
    df.to_excel(output_path, index=False)
    print(f"Zapisano plik z predykcją: {output_path}")

if __name__ == "__main__":
    run_prediction("../data/scalona_ocena.xlsx", "../data/predykcja_ml.xlsx")
