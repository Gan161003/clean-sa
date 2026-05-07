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
# TABLE DETECTION CONFIG
# =========================================================

IGNORE_WORDS = [
    "total",
    "grand total",
    "summary"
]

MIN_HEADER_MATCH = 2

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

# =====================================
# SAFE SERIES HANDLER
# =====================================

def safe_series(df, col):

    data = df[col]

    # If duplicate column names return DataFrame
    if isinstance(data, pd.DataFrame):
        data = data.iloc[:, 0]

    return data.astype(str)


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


# =========================================================
# FIND ALL TABLES
# =========================================================

def find_all_tables(df):

    tables = []

    rows, cols = df.shape

    for r in range(rows):

        for c in range(cols):

            matched_headers = {}

            for scan_c in range(c, min(c + 8, cols)):

                cell = clean_string(df.iat[r, scan_c])

                mapped = map_column(cell)

                if mapped:
                    matched_headers[mapped] = scan_c

            if len(matched_headers) >= MIN_HEADER_MATCH:

                tables.append({
                    "header_row": r,
                    "start_col": c,
                    "headers": matched_headers
                })

    # remove duplicates

    unique_tables = []
    seen = set()

    for t in tables:

        key = (t["header_row"], t["start_col"])

        if key not in seen:

            unique_tables.append(t)
            seen.add(key)

    return unique_tables


# =========================================================
# FIND TABLE END
# =========================================================

def find_table_end(df, start_row, start_col):

    blank_count = 0

    for r in range(start_row + 1, len(df)):

        row_data = df.iloc[r, start_col:start_col + 8]

        non_blank = row_data.notna().sum()

        if non_blank == 0:
            blank_count += 1
        else:
            blank_count = 0

        if blank_count >= 2:
            return r - 2

    return len(df) - 1


# =========================================================
# GET TABLE TITLE
# =========================================================

def get_table_title(df, header_row, start_col):

    check_rows = [
        header_row - 1,
        header_row - 2,
        header_row - 3
    ]

    for r in check_rows:

        if r < 0:
            continue

        val = clean_string(df.iat[r, start_col])

        if val != "" and "date" not in val:
            return str(df.iat[r, start_col]).strip()

    return ""

# =========================================================
# STANDARDIZE DATAFRAME
# =========================================================

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


# =========================================================
# CLEAN NUMERIC
# =========================================================

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
            safe_series(df, col)
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
            .str.strip()
        )

        df[col] = df[col].replace("", np.nan)

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

        st.write(f"📂 Processing: {uploaded_file.name}")

        unique_key = extract_unique_key(uploaded_file.name)

        try:

            excel_file = pd.ExcelFile(uploaded_file)

        except Exception as e:

            st.error(f"Cannot open file: {uploaded_file.name}")
            continue

        # =====================================================
        # LOOP SHEETS
        # =====================================================

        for sheet_name in excel_file.sheet_names:

            try:

                raw_df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    header=None
                )
                # Fix merged cells
                raw_df = raw_df.ffill()
                
                # Remove duplicate columns
                raw_df = raw_df.loc[:, ~raw_df.columns.duplicated()]

                raw_df = raw_df.dropna(
                    how="all"
                ).reset_index(drop=True)

                tables = find_all_tables(raw_df)

                if len(tables) == 0:

                    st.warning(
                        f"No valid table found in: {sheet_name}"
                    )

                    continue

                # =================================================
                # LOOP TABLES
                # =================================================

                for table in tables:

                    header_row = table["header_row"]
                    start_col = table["start_col"]

                    end_row = find_table_end(
                        raw_df,
                        header_row,
                        start_col
                    )

                    temp_df = raw_df.iloc[
                        header_row + 1:end_row + 1,
                        start_col:start_col + 8
                    ].copy()

                    actual_columns = []

                    for c in range(
                        start_col,
                        start_col + temp_df.shape[1]
                    ):

                        header_value = raw_df.iat[
                            header_row,
                            c
                        ]

                        actual_columns.append(header_value)

                    temp_df.columns = actual_columns

                    temp_df = temp_df.ffill()

                    temp_df = temp_df.dropna(
                        how="all"
                    )

                    # =============================================
                    # REMOVE TOTAL ROWS
                    # =============================================

                    mask = temp_df.apply(
                        lambda row: is_total_row(row),
                        axis=1
                    )

                    temp_df = temp_df[~mask]

                    # =============================================
                    # STANDARDIZE
                    # =============================================

                    temp_df = standardize_dataframe(temp_df)

                    temp_df = clean_numeric(temp_df)

                    # =============================================
                    # REMOVE EMPTY DATES
                    # =============================================

                    temp_df = temp_df[
                        safe_series(temp_df, "date")
                        .str.strip() != ""
                    ]

                    # =============================================
                    # SKIP EMPTY TABLE
                    # =============================================

                    if len(temp_df) == 0:
                        continue

                    # =============================================
                    # TABLE TITLE
                    # =============================================

                    table_title = get_table_title(
                        raw_df,
                        header_row,
                        start_col
                    )

                    temp_df["creative"] = table_title

                    # =============================================
                    # EXTRA COLUMNS
                    # =============================================

                    temp_df["unique_key"] = unique_key
                    temp_df["source_file"] = uploaded_file.name
                    temp_df["sheet_name"] = sheet_name

                    # =============================================
                    # FINAL COLUMN ORDER
                    # =============================================

                    final_cols = [
                        "unique_key",
                        "creative",
                        "date",
                        "impressions",
                        "clicks",
                        "views",
                        "spends",
                        "engagements",
                        "source_file",
                        "sheet_name"
                    ]

                    temp_df = temp_df[final_cols]

                    all_data.append(temp_df)

            except Exception as e:

                st.warning(
                    f"❌ Sheet Error: {sheet_name} | {str(e)}"
                )

        progress.progress(
            (idx + 1) / len(uploaded_files)
        )

    # =====================================================
    # FINAL OUTPUT
    # =====================================================

    if len(all_data) > 0:

        final_df = pd.concat(
            all_data,
            ignore_index=True
        )

        final_df = final_df.drop_duplicates()

        st.success("✅ Processing Completed")

        st.write(
            f"Total Rows Extracted: {len(final_df)}"
        )

        st.dataframe(final_df)

        # =================================================
        # DOWNLOAD EXCEL
        # =================================================

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

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

        st.error("❌ No Data Extracted")
