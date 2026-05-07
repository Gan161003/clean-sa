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
        "day"
    ],

    "impressions": [
        "impression",
        "impressions"
    ],

    "clicks": [
        "click",
        "clicks",
        "tap",
        "taps"
    ],

    "views": [
        "view",
        "views"
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


def clean_numeric(series):

    return (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace("", np.nan)
    )


# =========================================================
# FIND TABLES
# =========================================================

# def find_tables(df):

#     tables = []

#     rows, cols = df.shape

#     for r in range(rows):

#         for c in range(cols):

#             val = clean_string(
#                 df.iat[r, c]
#             )

#             if val != "date":
#                 continue

#             # CHECK NEXT 4 COLS
#             headers = []

#             for x in range(c, min(c + 5, cols)):

#                 cell = clean_string(
#                     df.iat[r + 1, x]
#                 )

#                 headers.append(cell)

#             # MUST HAVE IMPRESSION + CLICKS
#             has_imp = any(
#                 "impression" in h
#                 for h in headers
#             )

#             has_click = any(
#                 "click" in h or "tap" in h
#                 for h in headers
#             )

#             if has_imp and has_click:

#                 tables.append({
#                     "header_row": r,
#                     "start_col": c
#                 })

#     return tables




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

        val = df.iat[r, start_col]

        if pd.notna(val):

            val = str(val).strip()

            if (
                val != ""
                and val.lower() != "date"
            ):
                return val

    return ""


# =========================================================
# FIND TABLE END
# =========================================================

def find_table_end(
    df,
    start_row,
    start_col
):

    blank_count = 0

    for r in range(start_row + 2, len(df)):

        row_slice = df.iloc[
            r,
            start_col:start_col + 6
        ]

        non_blank = row_slice.notna().sum()

        if non_blank == 0:

            blank_count += 1

        else:

            blank_count = 0

        if blank_count >= 2:
            return r - 2

    return len(df) - 1


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

        except Exception as e:

            st.error(
                f"Cannot open file: "
                f"{uploaded_file.name}"
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

                raw_df = raw_df.dropna(
                    how="all"
                ).reset_index(drop=True)

                tables = find_tables(raw_df)

                if len(tables) == 0:
                    continue

                # =================================================
                # LOOP TABLES
                # =================================================

                for table in tables:

                    header_row = table[
                        "header_row"
                    ]

                    start_col = table[
                        "start_col"
                    ]

                    end_row = find_table_end(
                        raw_df,
                        header_row,
                        start_col
                    )

                    # =================================================
                    # EXTRACT DATA
                    # =================================================

                    temp_df = raw_df.iloc[
                        header_row + 2:end_row + 1,
                        start_col:start_col + 6
                    ].copy()

                    # =================================================
                    # HEADERS
                    # =================================================

                    actual_headers = []

                    for c in range(
                        start_col,
                        start_col + 5
                    ):

                        val = raw_df.iat[
                            header_row + 1,
                            c
                        ]

                        actual_headers.append(
                            str(val).strip()
                        )

                    # FIRST COL ALWAYS DATE
                    actual_headers[0] = "date"

                    temp_df.columns = actual_headers

                    # =================================================
                    # REMOVE TOTAL ROWS
                    # =================================================

                    temp_df = temp_df[
                        ~temp_df.apply(
                            is_total_row,
                            axis=1
                        )
                    ]

                    # =================================================
                    # STANDARDIZE
                    # =================================================

                    mapped_columns = {}

                    for col in temp_df.columns:

                        mapped = map_column(col)

                        if mapped:
                            mapped_columns[col] = mapped

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
                            temp_df[col] = np.nan

                    temp_df = temp_df[
                        required_cols
                    ]

                    # =================================================
                    # REMOVE EMPTY DATE
                    # =================================================

                    temp_df = temp_df[
                        temp_df["date"]
                        .notna()
                    ]

                    # =================================================
                    # DATE PARSE
                    # =================================================

                    temp_df["date"] = pd.to_datetime(
                        temp_df["date"],
                        errors="coerce"
                    )

                    temp_df = temp_df[
                        temp_df["date"].notna()
                    ]

                    if len(temp_df) == 0:
                        continue

                    # =================================================
                    # CLEAN NUMBERS
                    # =================================================

                    numeric_cols = [
                        "impressions",
                        "clicks",
                        "views",
                        "spends",
                        "engagements"
                    ]

                    for col in numeric_cols:

                        temp_df[col] = clean_numeric(
                            temp_df[col]
                        )

                    # =================================================
                    # TITLE
                    # =================================================

                    table_title = get_table_title(
                        raw_df,
                        header_row,
                        start_col
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

                    temp_df = temp_df.drop_duplicates()

                    all_data.append(temp_df)

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

        final_df = final_df.drop_duplicates()

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
