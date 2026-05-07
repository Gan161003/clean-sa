# =========================================================
# UNIVERSAL MEDIA REPORT CLEANER ENGINE (SEMI-AUTOMATION)
# =========================================================
#
# FINAL OUTPUT FORMAT:
# unique_key
# date
# impressions
# clicks
# views
# spends
# engagements
# source_file
# sheet_name
#
# =========================================================

import pandas as pd
import numpy as np
import os
import re
from openpyxl import load_workbook

# =========================================================
# CONFIG
# =========================================================

INPUT_FOLDER = "input_files"
OUTPUT_FILE = "Unified_Output.xlsx"

# =========================================================
# COLUMN MAPPING
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
        "video views",
        "view"
    ],

    "spends": [
        "spend",
        "spends",
        "cost",
        "media cost"
    ],

    "engagements": [
        "engagement",
        "engagements"
    ]
}

# =========================================================
# HELPER FUNCTIONS
# =========================================================

def clean_string(x):
    if pd.isna(x):
        return ""

    return str(x).strip().lower()


def extract_unique_key(filename):

    patterns = [
        r"(1ur-\d+)",
        r"(tur-\d+)",
        r"(ur-\d+)"
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

    total_keywords = [
        "total",
        "grand total",
        "summary"
    ]

    for keyword in total_keywords:

        if keyword in row_text:
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

    final_columns = [
        "date",
        "impressions",
        "clicks",
        "views",
        "spends",
        "engagements"
    ]

    for col in final_columns:

        if col not in df.columns:
            df[col] = ""

    return df[final_columns]


def clean_numeric_columns(df):

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
# MAIN FILE PROCESSOR
# =========================================================

all_data = []

files = [
    f for f in os.listdir(INPUT_FOLDER)
    if f.endswith((".xlsx", ".xls"))
]

for file in files:

    print("=" * 60)
    print(f"PROCESSING FILE: {file}")

    file_path = os.path.join(INPUT_FOLDER, file)

    unique_key = extract_unique_key(file)

    try:

        excel_file = pd.ExcelFile(file_path)

        for sheet_name in excel_file.sheet_names:

            print(f"Reading Sheet: {sheet_name}")

            try:

                raw_df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    header=None
                )

                # =================================================
                # FIND HEADER ROW
                # =================================================

                header_row = find_header_row(raw_df)

                if header_row is None:
                    print("No valid header found.")
                    continue

                print(f"Header Found At Row: {header_row}")

                # =================================================
                # RELOAD WITH HEADER
                # =================================================

                df = pd.read_excel(
                    file_path,
                    sheet_name=sheet_name,
                    header=header_row
                )

                # =================================================
                # CLEAN EMPTY ROWS
                # =================================================

                df = df.dropna(how="all")

                # =================================================
                # REMOVE TOTAL ROWS
                # =================================================

                mask = df.apply(
                    lambda row: is_total_row(row),
                    axis=1
                )

                df = df[~mask]

                # =================================================
                # STANDARDIZE
                # =================================================

                df = standardize_dataframe(df)

                # =================================================
                # CLEAN NUMBERS
                # =================================================

                df = clean_numeric_columns(df)

                # =================================================
                # REMOVE EMPTY DATE ROWS
                # =================================================

                df = df[
                    df["date"].astype(str).str.strip() != ""
                ]

                # =================================================
                # ADD METADATA
                # =================================================

                df["unique_key"] = unique_key
                df["source_file"] = file
                df["sheet_name"] = sheet_name

                # =================================================
                # FINAL COLUMN ORDER
                # =================================================

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

                # =================================================
                # APPEND
                # =================================================

                all_data.append(df)

                print(f"Rows Extracted: {len(df)}")

            except Exception as e:
                print(f"Sheet Error: {sheet_name}")
                print(str(e))

    except Exception as e:
        print(f"File Error: {file}")
        print(str(e))


# =========================================================
# FINAL OUTPUT
# =========================================================

if len(all_data) > 0:

    final_df = pd.concat(all_data, ignore_index=True)

    final_df.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print("=" * 60)
    print("UNIFIED OUTPUT CREATED SUCCESSFULLY")
    print(f"Total Rows: {len(final_df)}")
    print(f"Saved As: {OUTPUT_FILE}")

else:

    print("NO DATA EXTRACTED")
