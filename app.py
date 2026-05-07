import streamlit as st
import pandas as pd
import numpy as np
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
# CONFIG
# =========================================================

MIN_HEADER_MATCH = 3

FINAL_COLUMNS = [
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

# =========================================================
# HELPERS
# =========================================================

def clean_string(x):

    if pd.isna(x):
        return ""

    return str(x).strip().lower()


# =========================================================
# UNIQUE KEY
# =========================================================

def extract_unique_key(filename):

    patterns = [
        r"(1ur-[A-Za-z0-9]+)",
        r"(tur-[A-Za-z0-9]+)",
        r"(ur-[A-Za-z0-9]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            filename,
            re.IGNORECASE
        )

        if match:
            return match.group(1)

    return ""


# =========================================================
# COLUMN MAPPING
# =========================================================

def map_column(col_name):

    col_clean = clean_string(col_name)

    exact_map = {

        "date": "date",
        "day": "date",
        "time": "date",

        "impression": "impressions",
        "impressions": "impressions",
        "displays": "impressions",

        "click": "clicks",
        "clicks": "clicks",
        "tap": "clicks",
        "taps": "clicks",

        "views": "views",
        "video views": "views",

        "spend": "spends",
        "cost": "spends",

        "engagement": "engagements",
        "engagements": "engagements"
    }

    return exact_map.get(col_clean)


# =========================================================
# SAFE SERIES
# =========================================================

def safe_series(df, col):

    if col not in df.columns:
        return pd.Series(dtype=str)

    data = df[col]

    if isinstance(data, pd.DataFrame):
        data = data.iloc[:, 0]

    return data.astype(str)


# =========================================================
# REMOVE TOTAL ROWS
# =========================================================

def is_total_row(row):

    row_text = " ".join(
        [clean_string(x) for x in row]
    )

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
# FIND TABLES
# =========================================================

def find_all_tables(df):

    tables = []

    rows, cols = df.shape

    for r in range(rows):

        for c in range(cols):

            matched_headers = {}

            for scan_c in range(c, min(c + 8, cols)):

                cell = clean_string(
                    df.iat[r, scan_c]
                )

                mapped = map_column(cell)

                if mapped:
                    matched_headers[mapped] = scan_c

            if len(matched_headers) >= MIN_HEADER_MATCH:

                tables.append({
                    "header_row": r,
                    "start_col": c
                })

    # REMOVE DUPLICATES

    unique_tables = []
    seen = set()

    for t in tables:

        key = (
            t["header_row"],
            t["start_col"]
        )

        if key not in seen:

            unique_tables.append(t)
            seen.add(key)

    return sorted(
        unique_tables,
        key=lambda x: (
            x["header_row"],
            x["start_col"]
        )
    )


# =========================================================
# FIND TABLE END
# =========================================================

def find_table_end(df, start_row):

    blank_count = 0

    for r in range(start_row + 1, len(df)):

        row_data = df.iloc[r]

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

def get_table_title(
    df,
    header_row,
    start_col
):

    check_rows = [
        header_row - 1,
        header_row - 2,
        header_row - 3
    ]

    for r in check_rows:

        if r < 0:
            continue

        val = clean_string(
            df.iat[r, start_col]
        )

        if (
            val != ""
            and "date" not in val
            and "impression" not in val
            and "click" not in val
        ):

            return str(
                df.iat[r, start_col]
            ).strip()

    return ""


# =========================================================
# VALID TABLE
# =========================================================

def is_valid_table(df):

    if len(df) < 1:
        return False

    if "date" not in df.columns:
        return False

    valid_dates = pd.to_datetime(
        df["date"],
        errors="coerce"
    ).notna().sum()

    if valid_dates < 1:
        return False

    return True


# =========================================================
# STANDARDIZE
# =========================================================

def standardize_dataframe(df):

    mapped_columns = {}

    used = set()

    for col in df.columns:

        mapped = map_column(col)

        if mapped and mapped not in used:

            mapped_columns[col] = mapped
            used.add(mapped)

    df = df.rename(
        columns=mapped_columns
    )

    needed_cols = [
        "date",
        "impressions",
        "clicks",
        "views",
        "spends",
        "engagements"
    ]

    for col in needed_cols:

        if col not in df.columns:
            df[col] = ""

    return df[needed_cols]


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

        df[col] = df[col].replace(
            "",
            np.nan
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

        st.write(
            f"📂 Processing: {uploaded_file.name}"
        )

        unique_key = extract_unique_key(
            uploaded_file.name
        )

        try:

            excel_file = pd.ExcelFile(
                uploaded_file
            )

        except Exception:

            st.error(
                f"Cannot open file: {uploaded_file.name}"
            )

            continue

        # =====================================================
        # SHEETS
        # =====================================================

        for sheet_name in excel_file.sheet_names:

            try:

                raw_df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    header=None
                )

                # FIX MERGED CELLS
                raw_df = raw_df.ffill()

                # REMOVE BLANK ROWS
                raw_df = raw_df.dropna(
                    how="all"
                ).reset_index(drop=True)

                # FIND TABLES
                tables = find_all_tables(
                    raw_df
                )

                if len(tables) == 0:

                    st.warning(
                        f"No valid table found in: {sheet_name}"
                    )

                    continue

                # =================================================
                # TABLE LOOP
                # =================================================

                for i, table in enumerate(tables):

                    header_row = table["header_row"]
                    start_col = table["start_col"]

                    end_row = find_table_end(
                        raw_df,
                        header_row
                    )

                    # ============================================
                    # FIND NEXT TABLE COLUMN
                    # ============================================

                    next_table_col = raw_df.shape[1]

                    future_tables = sorted([
                        t["start_col"]
                        for t in tables
                        if (
                            t["header_row"] == header_row
                            and t["start_col"] > start_col
                        )
                    ])

                    if len(future_tables) > 0:
                        next_table_col = future_tables[0]

                    # ============================================
                    # EXTRACT TABLE
                    # ============================================

                    temp_df = raw_df.iloc[
                        header_row + 2:end_row + 1,
                        start_col:next_table_col
                    ].copy()

                    # REMOVE EMPTY COLUMNS
                    temp_df = temp_df.dropna(
                        axis=1,
                        how="all"
                    )

                    if temp_df.empty:
                        continue

                    # ============================================
                    # HEADERS
                    # ============================================


                    # ============================================
                    # MULTI HEADER EXTRACTION
                    # ============================================
                    
                    actual_columns = []
                    
                    used_headers = {}
                    
                    for idx_col, c in enumerate(
                        range(
                            start_col,
                            start_col + temp_df.shape[1]
                        )
                    ):
                    
                        top_header = clean_string(
                            raw_df.iat[header_row, c]
                        )
                    
                        second_header = clean_string(
                            raw_df.iat[header_row + 1, c]
                        )
                    
                        # USE SECOND HEADER IF EXISTS
                        if second_header != "":
                            header_value = second_header
                    
                        else:
                            header_value = top_header
                    
                        if header_value == "":
                            header_value = f"unknown_{idx_col}"
                    
                        # HANDLE DUPLICATES
                        if header_value in used_headers:
                    
                            used_headers[header_value] += 1
                    
                            header_value = (
                                f"{header_value}_"
                                f"{used_headers[header_value]}"
                            )
                    
                        else:
                    
                            used_headers[header_value] = 0
                    
                        actual_columns.append(header_value)
                    
                    # IMPORTANT
                    temp_df.columns = actual_columns

                    # FILL MERGED DATE CELLS
                    temp_df = temp_df.ffill(
                        axis=0
                    )

                    # REMOVE EMPTY ROWS
                    temp_df = temp_df.dropna(
                        how="all"
                    )

                    # REMOVE TOTAL ROWS
                    mask = temp_df.apply(
                        lambda row: is_total_row(row),
                        axis=1
                    )

                    temp_df = temp_df[~mask]

                    # STANDARDIZE
                    temp_df = standardize_dataframe(
                        temp_df
                    )

                    # CLEAN NUMBERS
                    temp_df = clean_numeric(
                        temp_df
                    )

                    # VALIDATE
                    if not is_valid_table(temp_df):
                        continue

                    # REMOVE EMPTY DATES
                    temp_df = temp_df[
                        safe_series(
                            temp_df,
                            "date"
                        ).str.strip() != ""
                    ]

                    if len(temp_df) == 0:
                        continue

                    # ============================================
                    # TABLE TITLE
                    # ============================================

                    table_title = get_table_title(
                        raw_df,
                        header_row,
                        start_col
                    )

                    temp_df["creative"] = (
                        table_title
                    )

                    # EXTRA FIELDS
                    temp_df["unique_key"] = (
                        unique_key
                    )

                    temp_df["source_file"] = (
                        uploaded_file.name
                    )

                    temp_df["sheet_name"] = (
                        sheet_name
                    )

                    # FINAL ORDER
                    temp_df = temp_df[
                        FINAL_COLUMNS
                    ]

                    # REMOVE DUPLICATES
                    temp_df = (
                        temp_df.drop_duplicates()
                    )

                    if len(temp_df) == 0:
                        continue

                    all_data.append(temp_df)

            except Exception as e:

                st.warning(
                    f"❌ Sheet Error: "
                    f"{sheet_name} | {str(e)}"
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

        final_df = (
            final_df.drop_duplicates()
        )

        st.success(
            "✅ Processing Completed"
        )

        st.write(
            f"Total Rows Extracted: "
            f"{len(final_df)}"
        )

        st.dataframe(final_df)

        # =================================================
        # DOWNLOAD
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
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    else:

        st.error(
            "❌ No Data Extracted"
        )
