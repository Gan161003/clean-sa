import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Universal Media Report Cleaner",
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
        "display",
        "displays"
    ],

    "clicks": [
        "click",
        "clicks",
        "tap",
        "taps"
    ],

    "views": [
        "view",
        "views",
        "video views"
    ],

    "spends": [
        "spend",
        "cost",
        "amount spent"
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

        match = re.search(
            pattern,
            filename,
            re.IGNORECASE
        )

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
# FIND ALL TABLES
# =========================================================
def find_all_tables(df):

    tables = []

    rows, cols = df.shape

    for r in range(rows):

        for c in range(cols):

            cell = clean_string(df.iat[r, c])

            # TABLE MUST START WITH DATE
            if "date" not in cell:
                continue

            matched_headers = {}

            for scan_c in range(c, min(c + 8, cols)):

                scan_cell = clean_string(
                    df.iat[r, scan_c]
                )

                mapped = map_column(scan_cell)

                if mapped:
                    matched_headers[mapped] = scan_c

            # MUST HAVE DATE + 1 METRIC
            metric_count = len(
                [
                    x for x in matched_headers
                    if x != "date"
                ]
            )

            if metric_count >= 1:

                tables.append({
                    "header_row": r,
                    "start_col": c,
                    "headers": matched_headers
                })

    # =========================================
    # REMOVE OVERLAPPING TABLES
    # =========================================

    final_tables = []

    used_positions = set()

    for t in tables:

        key = (
            t["header_row"],
            t["start_col"]
        )

        if key in used_positions:
            continue

        final_tables.append(t)

        # BLOCK NEARBY COLUMNS
        for i in range(
            t["start_col"],
            t["start_col"] + 5
        ):

            used_positions.add(
                (t["header_row"], i)
            )

    return final_tables

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

    for r in [
        header_row - 1,
        header_row - 2,
        header_row - 3
    ]:

        if r < 0:
            continue

        val = clean_string(
            df.iat[r, start_col]
        )

        if val != "" and "date" not in val:

            return str(
                df.iat[r, start_col]
            ).strip()

    return ""


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

        if col not in df.columns:
            df[col] = ""

        df[col] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("%", "", regex=False)
            .str.strip()
        )

        df[col] = df[col].replace(
            ["", "nan", "None"],
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
# PROCESS FILES
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

        except:

            st.error(
                f"Cannot open: {uploaded_file.name}"
            )

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

                # FIX MERGED CELLS
                raw_df = raw_df.ffill()

                raw_df = raw_df.dropna(
                    how="all"
                ).reset_index(drop=True)

                tables = find_all_tables(
                    raw_df
                )

                if len(tables) == 0:
                    continue

                # =================================================
                # LOOP TABLES
                # =================================================

                for i, table in enumerate(tables):

                    header_row = table[
                        "header_row"
                    ]

                    start_col = table[
                        "start_col"
                    ]



                    end_row = find_table_end(
                        raw_df,
                        header_row
                    )

                    # =================================================
                    # EXTRACT TABLE
                    # =================================================

                    # temp_df = raw_df.iloc[
                    #     header_row + 2:end_row + 1,
                    #     start_col:next_table_col
                    # ].copy()

                    TABLE_WIDTH = 5
                    
                    temp_df = raw_df.iloc[
                        header_row + 2:end_row + 1,
                        start_col:start_col + TABLE_WIDTH
                    ].copy()

                    temp_df = temp_df.dropna(
                        axis=1,
                        how="all"
                    )

                    if temp_df.empty:
                        continue

                    # =================================================
                    # ============================================
                    # MULTI HEADER EXTRACTION
                    # ============================================
                    
                    actual_columns = []
                    
                    used_headers = {}
                    
                    for idx_col in range(len(temp_df.columns)):
                    
                        c = temp_df.columns[idx_col]
                    
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
                    
                    # APPLY HEADERS
                    temp_df.columns = actual_columns
                    # =================================================
                    # REMOVE TOTAL ROWS
                    # =================================================

                    mask = temp_df.apply(
                        lambda row:
                        is_total_row(row),
                        axis=1
                    )

                    temp_df = temp_df[~mask]

                    # =================================================
                    # STANDARDIZE COLUMNS
                    # =================================================

                    mapped_columns = {}

                    used_cols = set()

                    for col in temp_df.columns:

                        mapped = map_column(col)

                        if mapped:

                            if mapped not in used_cols:

                                mapped_columns[
                                    col
                                ] = mapped

                                used_cols.add(
                                    mapped
                                )

                    temp_df = temp_df.rename(
                        columns=mapped_columns
                    )

                    # =================================================
                    # REQUIRED COLS
                    # =================================================

                    required_cols = [
                        "date",
                        "impressions",
                        "clicks",
                        "views",
                        "spends",
                        "engagements"
                    ]

                    for col in required_cols:

                        if col not in temp_df.columns:
                            temp_df[col] = ""

                    temp_df = temp_df[
                        required_cols
                    ]

                    # =================================================
                    # REMOVE EMPTY DATES
                    # =================================================

                    temp_df = temp_df[
                        temp_df["date"]
                        .astype(str)
                        .str.strip() != ""
                    ]

                    # =================================================
                    # CLEAN NUMERIC
                    # =================================================

                    temp_df = clean_numeric(
                        temp_df
                    )

                    # =================================================
                    # TABLE TITLE
                    # =================================================

                    table_title = (
                        get_table_title(
                            raw_df,
                            header_row,
                            start_col
                        )
                    )

                    # =================================================
                    # EXTRA COLS
                    # =================================================

                    temp_df["creative"] = (
                        table_title
                    )

                    temp_df["unique_key"] = (
                        unique_key
                    )

                    temp_df["source_file"] = (
                        uploaded_file.name
                    )

                    temp_df["sheet_name"] = (
                        sheet_name
                    )

                    # =================================================
                    # FINAL ORDER
                    # =================================================

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

                    temp_df = temp_df[
                        final_cols
                    ]

                    temp_df = (
                        temp_df
                        .drop_duplicates()
                    )

                    if len(temp_df) > 0:

                        all_data.append(
                            temp_df
                        )

            except Exception as e:

                st.warning(
                    f"❌ Sheet Error: "
                    f"{sheet_name} | {str(e)}"
                )

        progress.progress(
            (idx + 1)
            / len(uploaded_files)
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
            final_df
            .drop_duplicates()
        )

        st.success(
            "✅ Processing Completed"
        )

        st.write(
            f"Total Rows Extracted: "
            f"{len(final_df)}"
        )

        st.dataframe(
            final_df,
            use_container_width=True
        )

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
                "application/"
                "vnd.openxmlformats-"
                "officedocument."
                "spreadsheetml.sheet"
            )
        )

    else:

        st.error("❌ No Data Extracted")
