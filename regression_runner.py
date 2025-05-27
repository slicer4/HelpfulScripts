import pandas as pd
import statsmodels.api as sm
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import LabelEncoder
from sklearn.model_selection import train_test_split
import os

# === CONFIGURATION SECTION ===
CONFIG = {
    "file_path": "your_data.xlsx",  # Input file path
    "sheet_name": 0,                # Sheet index or name
    "nominal_columns": ["Control Function", "Automation Level", "Loss Type"],
    "quantitative_dependent": ["Cyber Maturity", "Incident Pattern Relevance"],
    "categorical_dependent": ["Control Function", "Automation Level", "Loss Type"],
    "independent_variable": "Success Rate Sum",
    "informational_columns": ["NIST ID"]
}
OUTPUT_FILE = "regression_output.xlsx"

# === HELPER FUNCTIONS ===
def run_linear_regression(df, y_col, x_cols):
    X = df[x_cols]
    X = sm.add_constant(X)
    y = df[y_col]
    model = sm.OLS(y, X).fit()
    return model

def run_multinomial_logistic_regression(df, y_col, x_cols):
    X = df[x_cols]
    y = LabelEncoder().fit_transform(df[y_col])
    model = LogisticRegression(multi_class="multinomial", solver="lbfgs", max_iter=1000)
    model.fit(X, y)
    return model, y

# === MAIN SCRIPT ===
def main():
    # Load data
    df = pd.read_excel(CONFIG["file_path"], sheet_name=CONFIG["sheet_name"])

    # One-hot encode nominal columns
    nominal_cols = CONFIG["nominal_columns"]
    df_encoded = pd.get_dummies(df, columns=nominal_cols, drop_first=True)

    # Identify base feature set
    independent = CONFIG["independent_variable"]
    all_columns = set(df_encoded.columns)
    dummies_created = all_columns - set(df.columns)
    features = [independent] + list(dummies_created)

    results = {}

    # Linear regressions
    for y in CONFIG["quantitative_dependent"]:
        model = run_linear_regression(df_encoded, y, features)
        results[f"Linear_{y}"] = model.summary().as_text()

    # Multinomial logistic regressions
    for y in CONFIG["categorical_dependent"]:
        model, y_encoded = run_multinomial_logistic_regression(df_encoded, y, features)
        coefs = pd.DataFrame(model.coef_, columns=features)
        coefs["Intercept"] = model.intercept_
        coefs["Class"] = model.classes_
        results[f"Logistic_{y}"] = coefs

    # Export results
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for name, result in results.items():
            if isinstance(result, str):  # statsmodels summary
                pd.DataFrame({"Summary": result.split('\n')}).to_excel(writer, sheet_name=name[:31], index=False)
            else:  # Coefficients table
                result.to_excel(writer, sheet_name=name[:31], index=False)

    print(f"Regression output written to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
