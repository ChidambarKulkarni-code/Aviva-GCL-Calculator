import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from io import BytesIO

# --------------------------------------------------
# Define exact paths for your two Excel files
# --------------------------------------------------
SINGLE_LIFE_FILE = Path("Aviva Final - Lap & HL - single life.xlsx")
JOINT_LIFE_FILE = Path("Aviva Group Credit Life Joint Life rates.xlsx")

# Configuration mapping selections to the correct file
PRODUCT_CONFIG = {
    "Home Loan": {
        "Single Life": {"file": SINGLE_LIFE_FILE},
        "Joint Life": {"file": JOINT_LIFE_FILE}
    },
    "Loan Against Property": {
        "Single Life": {"file": SINGLE_LIFE_FILE},
        "Joint Life": {"file": JOINT_LIFE_FILE}
    }
}

# Mapping exact tab names for sheet lookups
COVER_SHEET_MAP = {
    ("Home Loan", "Single Life", "Level Cover"): "Home Loan Level Cover",
    ("Home Loan", "Single Life", "Reducing Cover"): "Home Loan Reducing Cover",
    ("Home Loan", "Joint Life", "Level Cover"): "Home Loan Joint Level",
    ("Home Loan", "Joint Life", "Reducing Cover"): "Home Loan Joint Reducing",
    
    ("Loan Against Property", "Single Life", "Level Cover"): "Lap Level Cover",
    ("Loan Against Property", "Single Life", "Reducing Cover"): "Lap Reducing",
    ("Loan Against Property", "Joint Life", "Level Cover"): "Lap Joint Level",
    ("Loan Against Property", "Joint Life", "Reducing Cover"): "Lap Joint Reducing",
}

REQUIRED_COLUMNS = [
    "Customer Name",
    "Primary Age",
    "Secondary Age",      # Required, but can be blank/0 if Single Life
    "Loan Amount",
    "Tenure Months"
]

# --------------------------------------------------
# Page configuration
# --------------------------------------------------
st.set_page_config(
    page_title="Aviva GCL Premium Calculator",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("Aviva Group Credit Life (GCL) Premium Calculator")
st.caption(
    "Cross-file automated premium engine routing inputs seamlessly between Single and Joint Life rate books."
)
st.divider()

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------
@st.cache_data
def load_rate_table(file_path, sheet_name):
    try:
        raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found in {file_path.name}: {e}")

    header_matches = raw[raw.iloc[:, 0].astype(str).str.strip().eq("Entry Age")].index
    if len(header_matches) == 0:
        raise ValueError(f"'Entry Age' header row not found in sheet '{sheet_name}' inside {file_path.name}")

    header_row = header_matches[0]
    tenures = raw.iloc[header_row, 1:].dropna().astype(float).astype(int).tolist()

    table = raw.iloc[header_row + 1:, :len(tenures) + 1].dropna(how="all")
    table = table.rename(columns={0: "Entry Age"})
    table.columns = ["Entry Age"] + tenures
    table = table.dropna(subset=["Entry Age"])

    for tenure in tenures:
        table[tenure] = pd.to_numeric(table[tenure], errors="coerce")

    return table


def find_age_band(age, age_bands):
    for band in age_bands:
        text = str(band).strip()
        if "-" in text:
            low, high = text.split("-")
            if int(low) <= age <= int(high):
                return band
        else:
            try:
                if int(float(text)) == age:
                    return band
            except ValueError:
                continue
    return None


def get_rate(rate_table, age, tenure):
    tenure_cols = [col for col in rate_table.columns if col != "Entry Age"]
    if tenure not in tenure_cols:
        return None, f"Tenure {tenure} months not available."

    ages_available = rate_table["Entry Age"].tolist()
    if any("-" in str(age_value) for age_value in ages_available):
        matched_age = find_age_band(age, ages_available)
    else:
        matched_age = age if age in ages_available else None

    if matched_age is None:
        return None, f"Age {age} not found in rates."

    rate = rate_table.loc[rate_table["Entry Age"].astype(str) == str(matched_age), tenure].iloc[0]
    if pd.isna(rate):
        return None, "Rate is blank in matrix."

    return float(rate), "Rate found"


def calculate_premium(row, rate_table, life_type):
    try:
        age_1 = int(row["Primary Age"])
        loan_amount = float(row["Loan Amount"])
        tenure = int(row["Tenure Months"])
    except Exception:
        return pd.Series([np.nan, np.nan, "Invalid numeric input data"])

    # Look up rate from active table
    rate_1, remark_1 = get_rate(rate_table, age_1, tenure)
    if rate_1 is None:
        return pd.Series([np.nan, np.nan, f"Primary Lookup: {remark_1}"])

    final_rate = rate_1
    remark = "Calculated (Joint Sheet)" if life_type == "Joint Life" else "Calculated"

    premium = (loan_amount / 100000) * final_rate
    return pd.Series([final_rate, round(premium, 2), remark])


def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Aviva Premium Output")
    output.seek(0)
    return output


def read_uploaded_file(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file)


# --------------------------------------------------
# Sidebar Configurations
# --------------------------------------------------
st.sidebar.title("Aviva Underwriting Controls")

product = st.sidebar.selectbox("Select Product Type", ["Home Loan", "Loan Against Property"])
life_type = st.sidebar.selectbox("Select Structure",
