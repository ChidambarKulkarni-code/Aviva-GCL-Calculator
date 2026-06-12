import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from io import BytesIO

# --------------------------------------------------
# Define paths for BOTH Excel files
# --------------------------------------------------
HOME_LOAN_FILE = Path("Main GCL rates.xlsx")  # Adjust this if your HL file has a different name
LAP_FILE = Path("LAP_GCL_Rates.xlsx")          # Put your exact LAP file name here

# Centralized configuration mapping products to their respective files and tabs
PRODUCT_CONFIG = {
    "Home Loan": {
        "file": HOME_LOAN_FILE,
        "sheets": {
            ("Single Life", "Level Cover"): "Home Loan Level Cover",
            ("Single Life", "Reducing Cover"): "Home Loan Reducing Cover",
            ("Joint Life", "Level Cover"): "Home Loan Joint Level",
            ("Joint Life", "Reducing Cover"): "Home Loan Joint Reducing",
        }
    },
    "Loan Against Property": {
        "file": LAP_FILE,
        "sheets": {
            ("Single Life", "Level Cover"): "Lap Level Cover",
            ("Single Life", "Reducing Cover"): "Lap Reducing",
            ("Joint Life", "Level Cover"): "Lap Joint Level",
            ("Joint Life", "Reducing Cover"): "Lap Joint Reducing",
        }
    }
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
    "A multi-file portfolio premium computation engine for Single and Joint Life credit insurance coverage."
)

st.divider()

# --------------------------------------------------
# Helper Functions
# --------------------------------------------------
@st.cache_data
def load_rate_table(file_path, sheet_name):
    try:
        raw = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=None
        )
    except Exception as e:
        raise ValueError(f"Sheet '{sheet_name}' not found in {file_path.name}: {e}")

    header_matches = raw[
        raw.iloc[:, 0].astype(str).str.strip().eq("Entry Age")
    ].index

    if len(header_matches) == 0:
        raise ValueError(f"'Entry Age' header row not found in sheet '{sheet_name}' inside {file_path.name}")

    header_row = header_matches[0]

    tenures = (
        raw.iloc[header_row, 1:]
        .dropna()
        .astype(float)
        .astype(int)
        .tolist()
    )

    table = raw.iloc[
        header_row + 1:,
        :len(tenures) + 1
    ].dropna(how="all")

    table = table.rename(columns={0: "Entry Age"})
    table.columns = ["Entry Age"] + tenures
    table = table.dropna(subset=["Entry Age"])

    for tenure in tenures:
        table[tenure] = pd.to_numeric(
            table[tenure],
            errors="coerce"
        )

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

    rate = rate_table.loc[
        rate_table["Entry Age"].astype(str) == str(matched_age),
        tenure
    ].iloc[0]

    if pd.isna(rate):
        return None, "Rate is blank in matrix."

    return float(rate), "Rate found"


def calculate_premium(row, rate_table, life_type, sheet_name, joint_factor=1.7):
    try:
        age_1 = int(row["Primary Age"])
        loan_amount = float(row["Loan Amount"])
        tenure = int(row["Tenure Months"])
    except Exception:
        return pd.Series([np.nan, np.nan, "Invalid numeric input data"])

    # Look up rate for primary member
    rate_1, remark_1 = get_rate(rate_table, age_1, tenure)
    if rate_1 is None:
        return pd.Series([np.nan, np.nan, f"Primary: {remark_1}"])

    # Process pricing based on Life Type selection
    if life_type == "Joint Life":
        try:
            # Check if using separate dedicated Joint Sheets or if blending individual rates
            if "Joint" not in sheet_name:
                if pd.isna(row["Secondary Age"]) or str(row["Secondary Age"]).strip() == "":
                    return pd.Series([np.nan, np.nan, "Missing Secondary Age for Joint computation"])
                
                age_2 = int(row["Secondary Age"])
                rate_2, remark_2 = get_rate(rate_table, age_2, tenure)
                if rate_2 is None:
                    return pd.Series([np.nan, np.nan, f"Secondary: {remark_2}"])
                
                # Dynamic formula: Blended Joint Rate
                final_rate = (rate_1 + rate_2) * (joint_factor / 2)
                remark = "Calculated (Joint Blended)"
            else:
                final_rate = rate_1
                remark = "Calculated (Joint Sheet)"
        except Exception:
            return pd.Series([np.nan, np.nan, "Error processing Secondary Age"])
    else:
        final_rate = rate_1
        remark = "Calculated"

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

product = st.sidebar.selectbox(
    "Select Product Type",
    ["Home Loan", "Loan Against Property"]
)

life_type = st.sidebar.selectbox(
    "Select Structure",
    ["Single Life", "Joint Life"]
)

cover_type = st.sidebar.selectbox(
    "Select Cover Type",
    ["Level Cover", "Reducing Cover"]
)

# Route execution parameters to the chosen file and target sheet
active_file = PRODUCT_CONFIG[product]["file"]
sheet_name = PRODUCT_CONFIG[product]["sheets"].get((life_type, cover_type))

# Dynamic File-Availability Guardrail
if not active_file.exists():
    st.error(
        f"Critical Error: Base rate file missing. Please ensure **{active_file.name}** is placed in the root folder."
    )
    st.stop()

if sheet_name:
    st.sidebar.success(f"Linked File: {active_file.name}")
    st.sidebar.text(f"Active Tab: {sheet_name}")
else:
    st.sidebar.error("Selected matrix pairing is not mapped.")
    st.stop()

# Multiplier option if computing Joint risk using custom single rate factor formulas
joint_factor = 1.7
if life_type == "Joint Life" and "Joint" not in sheet_name:
    joint_factor = st.sidebar.slider(
        "Joint Life Multiplier Factor", 
        min_value=1.0, max_value=2.0, value=1.7, step=0.05,
        help="Adjusts premium scale if using standard single-sheet calculations for dual covers."
    )


# --------------------------------------------------
# Main Layout UI
# --------------------------------------------------
left_col, right_col = st.columns([2, 1])

with left_col:
    st.subheader("Upload Customer Portfolio File")
    uploaded_file = st.file_uploader(
        "Supported formats: .xlsx, .xls, .csv",
        type=["xlsx", "xls", "csv"]
    )

with right_col:
    st.subheader("Required Portfolio Format")
    st.write("Your spreadsheet headers must match exactly:")
    for col in REQUIRED_COLUMNS:
        st.markdown(f"- `{col}`")
    st.info(
        "Note: Internal indicators like Policy No, Branch codes, or RM names are safely preserved in output downloads."
    )

# --------------------------------------------------
# Execution & Computation Engine
# --------------------------------------------------
if uploaded_file is not None:
    try:
        customer_df = read_uploaded_file(uploaded_file)
    except Exception as e:
        st.error(f"Error accessing spreadsheet data structure: {e}")
        st.stop()

    st.divider()
    st.subheader("Data Input Preview (First 20 Columns)")
    st.dataframe(customer_df.head(20), use_container_width=True)

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in customer_df.columns]

    if missing_cols:
        st.error(f"Execution halted. Missing mandatory columns: {', '.join(missing_cols)}")
    else:
        st.success("Structure validation passed successfully.")

        c1, c2, c3 = st.columns(3)
        c1.metric("Uploaded Insured Headcount", len(customer_df))
        c2.metric("Target Segment", product)
        c3.metric("LOB Structure", f"{life_type} ({cover_type})")

        st.divider()

        if st.button("Run Portfolio Premium Calculation", type="primary"):
            try:
                rate_table = load_rate_table(active_file, sheet_name)
            except Exception as e:
                st.error(str(e))
                st.stop()

            output_df = customer_df.copy()

            # Execute underwriting matrix transformation 
            output_df[[
                "Rate Per Lakh",
                "Calculated Premium",
                "Remarks"
            ]] = output_df.apply(
                lambda row: calculate_premium(row, rate_table, life_type, sheet_name, joint_factor),
                axis=1
            )

            total_records = len(output_df)
            calculated_records = (output_df["Remarks"].str.contains("Calculated", na=False)).sum()
            failed_records = total_records - calculated_records
            total_premium = output_df["Calculated Premium"].sum(skipna=True)

            st.success("Calculations complete.")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Rows", total_records)
            m2.metric("Passed Items", calculated_records)
            m3.metric("Failed Exceptions", failed_records)
            m4.metric("GWP Premium Portfolio Total", f"₹{total_premium:,.2f}")

            st.divider()
            st.subheader("Calculated Premium Matrix Output")
            st.dataframe(output_df, use_container_width=True)

            if failed_records > 0:
                st.warning("Exceptions flagged. Verify age boundaries or tenure constraints in failed rows.")

            excel_file = convert_df_to_excel(output_df)
            st.download_button(
                label="Download Calculated Premium Excel",
                data=excel_file,
                file_name=f"aviva_gcl_premium_{product.lower().replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Awaiting file upload profile parameters to generate rates.")
