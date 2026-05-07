import streamlit as st
import pandas as pd
import numpy as np
import os
import re
from io import BytesIO

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Media Report Cleaner",
    layout="wide"
)

st.title("📊 Universal Media Report Cleaner")

# =========================================================
# COLUMN MAP
# =========================================================

COLUMN_MAP = {

    "date": [
        "date",
        "day",
        "time"
    ],

    "impressions": [
        "impression",
        "impressions",
        "displays",
        "total impressions"
    ],

    "clicks": [
        "click",
        "clicks",
        "tap",
        "taps",
        "click-throughs",
        "total taps"
    ],

    "views": [
        "views",
        "video views"
    ],

    "spends": [
        "spend",
        "cost"
    ],

    "engagements": [
        "engagement",
        "engagements"
    ]
}


# =========================================================
# HELPERS
# =========================================================

def clean_string(x):

    if pd.isna(x):
        return ""

    return str(x).strip().lower()


def extract_unique_key(filename):

    patterns = [
        r"(1ur-[A-Za-z0-9]+)",
        r"(tur-[A-Za-z0-9]+)",
        r"(ur-[A-Za-z0-9]+)"
    ]

    for pattern in patterns:

        match = re.search(pattern, filename, re.IGNORECASE)

        if match:
            return match.group(1)

    return ""


def map_column(col_name):

    col_clean = clean_string(col_name)

    for standard_col, aliases in COLUMN_MAP.items():

        for alias in aliases:

            if alias in col_clean:
                return standard_col

    return None


def is_total_row(row):

    row_text = " ".join([clean_string(x) for x in row])

    keywords = [
        "total",
        "grand total",
        "summary"
    ]

    for k in keywords:

        if k in row_text:
            return True

    return False


def find_header_row(df):

    best_row = None
    best_score = 0

    for i in range(min(20, len(df))):

        row = df.iloc[i].tolist()

        score = 0

        for cell in row:

            cell_clean = clean_string(cell)

            for aliases in COLUMN_MAP.values():

                for alias in aliases:

                    if alias in cell_clean:
                        score += 1

        if score > best_score:

            best_score = score
            best_row = i

    return best_row


def standardize_dataframe(df):

    mapped_columns = {}

    for col in df.columns:

        mapped = map_column(col)

        if mapped:
            mapped_columns[col] = mapped

    df = df.rename(columns=mapped_columns)

    final_cols = [
        "date",
        "impressions",
        "clicks",
        "views",
        "spends",
        "engagements"
    ]

    for col in final_cols:

        if col not in df.columns:
            df[col] = ""

    return df[final_cols]


def clean_numeric(df):

    numeric_cols = [
        "impressions",
        "clicks",
        "views",
        "spends",
        "engagements"
    ]

    for col in numeric_cols:

        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
            .str.strip()
        )

    return df


# =========================================================
# FILE UPLOADER
# =========================================================

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# =========================================================
# PROCESS
# =========================================================

if uploaded_files:

    all_data = []

    progress = st.progress(0)

    for idx, uploaded_file in enumerate(uploaded_files):

        st.write(f"Processing: {uploaded_file.name}")

        unique_key = extract_unique_key(uploaded_file.name)

        excel_file = pd.ExcelFile(uploaded_file)

        for sheet_name in excel_file.sheet_names:

            try:

                raw_df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    header=None
                )

                header_row = find_header_row(raw_df)

                if header_row is None:
                    continue

                df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    header=header_row
                )
                df = df.ffill()

                df = df.dropna(how="all")

                mask = df.apply(
                    lambda row: is_total_row(row),
                    axis=1
                )

                df = df[~mask]

                df = standardize_dataframe(df)

                df = clean_numeric(df)

                df = df[
                    df["date"].astype(str).str.strip() != ""
                ]

                df["unique_key"] = unique_key
                df["source_file"] = uploaded_file.name
                df["sheet_name"] = sheet_name

                final_cols = [
                    "unique_key",
                    "date",
                    "impressions",
                    "clicks",
                    "views",
                    "spends",
                    "engagements",
                    "source_file",
                    "sheet_name"
                ]

                df = df[final_cols]

                all_data.append(df)

            except Exception as e:

                st.warning(
                    f"Sheet Error: {sheet_name} | {str(e)}"
                )

        progress.progress((idx + 1) / len(uploaded_files))

    # =====================================================
    # FINAL OUTPUT
    # =====================================================

    if len(all_data) > 0:

        final_df = pd.concat(all_data, ignore_index=True)

        st.success("Processing Completed")

        st.dataframe(final_df)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            final_df.to_excel(
                writer,
                index=False,
                sheet_name="Unified_Data"
            )

        st.download_button(
            label="📥 Download Unified Output",
            data=output.getvalue(),
            file_name="Unified_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:

        st.error("No Data Extracted")
