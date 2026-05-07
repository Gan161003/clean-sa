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
        "view",
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
# DETECT FILE TYPE
# =========================================================

def detect_file_type(df):

    rows, cols = df.shape

    horizontal_score = 0

    for r in range(min(15, rows)):

        row_values = [
            clean_string(x)
            for x in df.iloc[r].tolist()
        ]

        date_count = row_values.count("date")

        if date_count >= 2:
            horizontal_score += 5

    for r in range(min(15, rows)):

        row_values = " ".join([
            clean_string(x)
            for x in df.iloc[r].tolist()
        ])

        if row_values.count("impression") >= 2:
            horizontal_score += 3

    if horizontal_score >= 5:
        return "horizontal"

    return "vertical"


# =========================================================
# VERTICAL PARSER
# =========================================================

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
            df[col] = np.nan

    return df[final_cols]


def process_vertical_sheet(
    raw_df,
    file_name,
    sheet_name,
    unique_key
):

    all_tables = []

    header_row = find_header_row(raw_df)

    if header_row is None:
        return None

    df = pd.DataFrame(
        raw_df.iloc[header_row + 1:].values,
        columns=raw_df.iloc[header_row]
    )

    # IMPORTANT
    df = df.ffill()

    df = df.dropna(how="all")

    mask = df.apply(
        lambda row: is_total_row(row),
        axis=1
    )

    df = df[~mask]

    df = standardize_dataframe(df)

    df = df[
        df["date"].astype(str).str.strip() != ""
    ]

    # DATE
    df["date"] = pd.to_datetime(
        df["date"],
        errors="coerce"
    )

    df = df[df["date"].notna()]

    # NUMERIC
    numeric_cols = [
        "impressions",
        "clicks",
        "views",
        "spends",
        "engagements"
    ]

    for col in numeric_cols:
        df[col] = clean_numeric(df[col])

    df["creative"] = sheet_name
    df["unique_key"] = unique_key
    df["source_file"] = file_name
    df["sheet_name"] = sheet_name

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

    df = df[final_cols]

    all_tables.append(df)

    if len(all_tables) == 0:
        return None

    return pd.concat(
        all_tables,
        ignore_index=True
    )


# =========================================================
# HORIZONTAL PARSER
# =========================================================

def find_horizontal_tables(df):

    tables = []

    rows, cols = df.shape

    for r in range(rows - 1):

        for c in range(cols):

            current = clean_string(
                df.iat[r, c]
            )

            if current != "date":
                continue

            found_imp = False
            found_click = False

            for scan_c in range(
                c,
                min(c + 6, cols)
            ):

                val1 = clean_string(
                    df.iat[r, scan_c]
                )

                val2 = clean_string(
                    df.iat[r + 1, scan_c]
                )

                combined = val1 + " " + val2

                if "impression" in combined:
                    found_imp = True

                if (
                    "click" in combined
                    or
                    "tap" in combined
                ):
                    found_click = True

            if found_imp and found_click:

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

    return unique_tables


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


def process_horizontal_sheet(
    raw_df,
    file_name,
    sheet_name,
    unique_key
):

    all_tables = []

    tables = find_horizontal_tables(raw_df)

    processed = set()

    for table in tables:

        header_row = table["header_row"]
        start_col = table["start_col"]

        # AVOID DUPLICATE TABLES
        table_key = (
            header_row,
            start_col
        )

        if table_key in processed:
            continue

        processed.add(table_key)

        end_row = find_table_end(
            raw_df,
            header_row,
            start_col
        )

        temp_df = raw_df.iloc[
            header_row + 2:end_row + 1,
            start_col:start_col + 6
        ].copy()

        if temp_df.empty:
            continue

        actual_headers = []

        total_cols = temp_df.shape[1]

        for c in range(total_cols):

            header_val = raw_df.iat[
                header_row + 1,
                start_col + c
            ]

            actual_headers.append(
                str(header_val).strip()
            )

        actual_headers[0] = "date"

        temp_df.columns = actual_headers

        # ONLY DATE FFILL
        temp_df.iloc[:, 0] = (
            temp_df.iloc[:, 0].ffill()
        )

        # REMOVE TOTAL
        temp_df = temp_df[
            ~temp_df.apply(
                is_total_row,
                axis=1
            )
        ]

        mapped_columns = {}

        for col in temp_df.columns:

            mapped = map_column(col)

            if mapped:
                mapped_columns[col] = mapped

        temp_df = temp_df.rename(
            columns=mapped_columns
        )

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

        temp_df = temp_df[required_cols]

        # DATE
        temp_df["date"] = pd.to_datetime(
            temp_df["date"],
            errors="coerce"
        )

        temp_df = temp_df[
            temp_df["date"].notna()
        ]

        if len(temp_df) == 0:
            continue

        # NUMERIC
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

        table_title = get_table_title(
            raw_df,
            header_row,
            start_col
        )

        temp_df["creative"] = table_title
        temp_df["unique_key"] = unique_key
        temp_df["source_file"] = file_name
        temp_df["sheet_name"] = sheet_name

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

        temp_df = temp_df.drop_duplicates()

        all_tables.append(temp_df)

    if len(all_tables) == 0:
        return None

    return pd.concat(
        all_tables,
        ignore_index=True
    )


# =========================================================
# FILE UPLOADER
# =========================================================

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# =========================================================
# MAIN PROCESS
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
        # SHEETS
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

                file_type = detect_file_type(
                    raw_df
                )

                st.write(
                    f"{sheet_name} → {file_type}"
                )

                # =============================================
                # HORIZONTAL
                # =============================================

                if file_type == "horizontal":

                    result = process_horizontal_sheet(
                        raw_df,
                        uploaded_file.name,
                        sheet_name,
                        unique_key
                    )

                # =============================================
                # VERTICAL
                # =============================================

                else:

                    result = process_vertical_sheet(
                        raw_df,
                        uploaded_file.name,
                        sheet_name,
                        unique_key
                    )

                if (
                    result is not None
                    and
                    len(result) > 0
                ):
                    all_data.append(result)

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
