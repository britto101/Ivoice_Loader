import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import html

from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit.components.v1 as components
from openpyxl import Workbook


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Bill / Invoice Search",
    page_icon="🔎",
    layout="wide"
)


# ============================================================
# PAGE TITLE
# ============================================================

st.title("🔎 Bill / Invoice Search")

st.caption(
    "Search exact Bill / Invoice numbers across uploaded Excel files."
)


# ============================================================
# SETTINGS
# ============================================================

# Number of rows above the Bill/Invoice header to search
# for missing column names.
HEADER_LOOKBACK = 5


# ============================================================
# BASIC HELPERS
# ============================================================

def is_blank(value):
    """
    Check whether an Excel cell is blank.
    """

    if value is None:
        return True

    try:

        if pd.isna(value):
            return True

    except Exception:
        pass

    return str(value).strip() == ""


# ============================================================
# NORMALIZE HEADER
#
# Used ONLY to detect Bill / Invoice headers.
#
# Original Excel header is never changed.
# ============================================================

def normalize_header(value):

    if is_blank(value):
        return ""

    value = unicodedata.normalize(
        "NFKC",
        str(value)
    )

    value = value.strip().lower()

    # Convert separators to spaces
    value = re.sub(
        r"[_\-/\\.:]+",
        " ",
        value
    )

    # Remove punctuation
    value = re.sub(
        r"[^\w\s]",
        " ",
        value
    )

    # Multiple spaces -> one
    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


# ============================================================
# DETECT BILL / INVOICE HEADER
# ============================================================

def is_bill_invoice_header(value):

    header = normalize_header(value)

    if not header:
        return False

    # Common exact formats
    exact_headers = {

        "bill no",
        "bil no",

        "bill number",
        "bil number",

        "bill num",
        "bil num",

        "invoice no",
        "invoice number",
        "invoice num",

        "inv no",
        "inv number",
        "inv num",

        "bill ref no",
        "bill ref number",

        "invoice ref no",
        "invoice ref number",

        "bill reference no",
        "bill reference number",

        "invoice reference no",
        "invoice reference number"
    }

    if header in exact_headers:
        return True

    # Word-based detection
    words = set(
        header.split()
    )

    bill_words = {
        "bill",
        "bil",
        "invoice",
        "inv"
    }

    number_words = {
        "no",
        "number",
        "num"
    }

    if (
        words.intersection(bill_words)
        and words.intersection(number_words)
    ):
        return True

    return False


# ============================================================
# NORMALIZE SEARCH VALUE
#
# Used ONLY for comparison.
#
# The original Excel value remains unchanged in results.
# ============================================================

def normalize_search_value(value):

    if is_blank(value):
        return ""

    # Excel number stored as float
    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

    if isinstance(value, int):
        return str(value)

    return str(value).strip()


# ============================================================
# CHECK EMPTY ROW
# ============================================================

def is_empty_row(row):

    for value in row:

        if not is_blank(value):
            return False

    return True


# ============================================================
# FIND ALL BILL / INVOICE HEADERS
# ============================================================

def find_all_bill_headers(df):

    found = []

    for row_number in range(
        len(df)
    ):

        row = df.iloc[
            row_number
        ]

        for column_number, value in enumerate(
            row
        ):

            if is_bill_invoice_header(
                value
            ):

                found.append({

                    "row":
                        row_number,

                    "column":
                        column_number,

                    "value":
                        value
                })

    return found


# ============================================================
# FIND END OF BILL / INVOICE TABLE
#
# Table ends when:
#
# 1. Another Bill/Invoice header begins
# OR
# 2. Two completely blank rows are found
# ============================================================

def find_table_end(
    df,
    start_row,
    next_header_row=None
):

    total_rows = len(df)

    if next_header_row is not None:

        limit = min(
            next_header_row,
            total_rows
        )

    else:

        limit = total_rows

    empty_rows = 0

    for row_number in range(
        start_row + 1,
        limit
    ):

        row = df.iloc[
            row_number
        ]

        if is_empty_row(row):

            empty_rows += 1

            if empty_rows >= 2:

                return row_number

        else:

            empty_rows = 0

    return limit


# ============================================================
# FIND COMPLETE TABLE WIDTH
#
# IMPORTANT:
#
# We do NOT rely only on the header row.
#
# We inspect the actual data rows too.
#
# This ensures columns after SHOP NAME are not lost.
# ============================================================

def find_table_width(
    df,
    header_row,
    table_end,
    bill_column
):

    last_column = bill_column

    # --------------------------------------------------------
    # Check header row
    # --------------------------------------------------------

    header = df.iloc[
        header_row
    ]

    for column_number, value in enumerate(
        header
    ):

        if not is_blank(value):

            last_column = max(
                last_column,
                column_number
            )

    # --------------------------------------------------------
    # Check every data row
    # --------------------------------------------------------

    for row_number in range(
        header_row + 1,
        table_end
    ):

        row = df.iloc[
            row_number
        ]

        for column_number, value in enumerate(
            row
        ):

            if not is_blank(value):

                last_column = max(
                    last_column,
                    column_number
                )

    return last_column + 1


# ============================================================
# FIND ACTUAL COLUMN HEADERS
#
# Example:
#
# Row 369:
#                         DVR | CAMERA | HDD | NVR
#
# Row 370:
# SNO | BILL NUMBER | SHOP ID | SHOP NAME | blank | blank
#
# Row 371:
# 1   | GST/...     | 17128   | ...       | 3970  | 68
#
# The function uses row 370 where available.
# For blank header cells it searches upward.
# ============================================================

def find_actual_headers(
    df,
    header_row,
    table_width
):

    final_headers = []

    for column_number in range(
        table_width
    ):

        # ----------------------------------------------------
        # First check Bill/Invoice header row
        # ----------------------------------------------------

        current_value = df.iat[
            header_row,
            column_number
        ]

        if not is_blank(
            current_value
        ):

            final_headers.append(
                current_value
            )

            continue

        # ----------------------------------------------------
        # Header cell is blank.
        #
        # Search upward for actual column name.
        # ----------------------------------------------------

        found_header = None

        start_row = max(
            0,
            header_row - HEADER_LOOKBACK
        )

        for previous_row in range(
            header_row - 1,
            start_row - 1,
            -1
        ):

            value = df.iat[
                previous_row,
                column_number
            ]

            if not is_blank(
                value
            ):

                found_header = value

                break

        final_headers.append(
            found_header
        )

    return final_headers


# ============================================================
# PROCESS ONE BILL / INVOICE TABLE
# ============================================================

def process_table(
    df,
    header_info,
    next_header_row=None
):

    header_row = header_info[
        "row"
    ]

    bill_column = header_info[
        "column"
    ]

    bill_header = header_info[
        "value"
    ]

    # --------------------------------------------------------
    # Find end
    # --------------------------------------------------------

    table_end = find_table_end(
        df,
        header_row,
        next_header_row
    )

    # --------------------------------------------------------
    # Find complete width
    # --------------------------------------------------------

    table_width = find_table_width(
        df,
        header_row,
        table_end,
        bill_column
    )

    # --------------------------------------------------------
    # Find actual column headers
    # --------------------------------------------------------

    headers = find_actual_headers(
        df,
        header_row,
        table_width
    )

    # --------------------------------------------------------
    # Data begins after Bill/Invoice header row
    # --------------------------------------------------------

    data_start = header_row + 1

    if data_start >= table_end:

        return None

    table_data = df.iloc[
        data_start:table_end,
        :table_width
    ]

    # --------------------------------------------------------
    # Build fast search index
    # --------------------------------------------------------

    search_map = defaultdict(list)

    for row in table_data.itertuples(
        index=False,
        name=None
    ):

        row = list(row)

        # Ensure complete width
        if len(row) < table_width:

            row.extend(
                [None] *
                (
                    table_width - len(row)
                )
            )

        # ----------------------------------------------------
        # Bill/Invoice column
        # ----------------------------------------------------

        if bill_column >= len(row):
            continue

        bill_value = row[
            bill_column
        ]

        search_value = normalize_search_value(
            bill_value
        )

        if not search_value:
            continue

        # Store the complete original row
        search_map[
            search_value
        ].append(
            row
        )

    return {

        "filename":
            None,

        "sheet":
            None,

        "header_row":
            header_row + 1,

        "bill_column":
            bill_column + 1,

        "bill_header":
            bill_header,

        "headers":
            headers,

        "search_map":
            dict(search_map)
    }


# ============================================================
# READ ONE EXCEL FILE
# ============================================================

def read_excel_file(
    filename,
    file_bytes
):

    result = {

        "filename":
            filename,

        "tables":
            [],

        "error":
            None
    }

    try:

        # ----------------------------------------------------
        # Read ALL sheets once
        # ----------------------------------------------------

        sheets = pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=None,
            header=None,
            dtype=object,
            engine="openpyxl"
        )

        # ----------------------------------------------------
        # Process every sheet
        # ----------------------------------------------------

        for sheet_name, df in sheets.items():

            if df is None or df.empty:
                continue

            # ------------------------------------------------
            # Remove completely empty rows ONLY.
            #
            # DO NOT remove empty columns.
            # ------------------------------------------------

            df = df.dropna(
                axis=0,
                how="all"
            ).reset_index(
                drop=True
            )

            if df.empty:
                continue

            # ------------------------------------------------
            # Find Bill/Invoice headers
            # ------------------------------------------------

            bill_headers = find_all_bill_headers(
                df
            )

            if not bill_headers:
                continue

            # ------------------------------------------------
            # Process every table
            # ------------------------------------------------

            for index, header_info in enumerate(
                bill_headers
            ):

                next_header_row = None

                if (
                    index + 1
                    < len(bill_headers)
                ):

                    next_header_row = (
                        bill_headers[
                            index + 1
                        ]["row"]
                    )

                table = process_table(
                    df,
                    header_info,
                    next_header_row
                )

                if table is None:
                    continue

                table["filename"] = filename

                table["sheet"] = sheet_name

                result["tables"].append(
                    table
                )

    except Exception as e:

        result["error"] = str(e)

    return result


# ============================================================
# BUILD SEARCH INDEX
#
# Cached so Excel is not reread every time the user searches.
# ============================================================

@st.cache_data(
    show_spinner=False,
    max_entries=5
)
def build_search_index(
    file_data
):

    results = []

    workers = min(
        4,
        max(
            1,
            len(file_data)
        )
    )

    with ThreadPoolExecutor(
        max_workers=workers
    ) as executor:

        futures = {}

        for filename, file_bytes in file_data:

            future = executor.submit(
                read_excel_file,
                filename,
                file_bytes
            )

            futures[future] = filename

        for future in as_completed(
            futures
        ):

            results.append(
                future.result()
            )

    # Keep predictable order
    results.sort(
        key=lambda x: x["filename"]
    )

    return results


# ============================================================
# SEARCH
#
# ONLY detected Bill/Invoice columns are searched.
#
# EXACT MATCH.
# ============================================================

def search_excel(
    indexed_files,
    search_value
):

    search_value = normalize_search_value(
        search_value
    )

    results = []

    for file_data in indexed_files:

        for table in file_data["tables"]:

            matching_rows = table[
                "search_map"
            ].get(
                search_value,
                []
            )

            for row in matching_rows:

                results.append({

                    "filename":
                        table["filename"],

                    "sheet":
                        table["sheet"],

                    "header":
                        table["bill_header"],

                    "header_row":
                        table["header_row"],

                    "bill_column":
                        table["bill_column"],

                    "headers":
                        table["headers"],

                    "row":
                        row
                })

    return results


# ============================================================
# SHOW DETECTED BILL / INVOICE HEADERS
# ============================================================

def show_detected_headers(
    indexed_files
):

    detected = []

    for file_data in indexed_files:

        for table in file_data["tables"]:

            detected.append({

                "File":
                    table["filename"],

                "Sheet":
                    table["sheet"],

                "Header":
                    table["bill_header"],

                "Excel Row":
                    table["header_row"],

                "Excel Column":
                    table["bill_column"]
            })

    if detected:

        st.dataframe(
            pd.DataFrame(
                detected
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# HTML VALUE
# ============================================================

def html_value(value):

    if value is None:
        return ""

    try:

        if pd.isna(value):
            return ""

    except Exception:
        pass

    return html.escape(
        str(value)
    )


# ============================================================
# DISPLAY RESULT TABLE
#
# Uses HTML because Excel may contain:
#
# - Duplicate headers
# - Blank headers
#
# st.dataframe() / PyArrow rejects duplicate names.
# ============================================================

def display_excel_table(
    headers,
    rows
):

    headers = list(
        headers
    )

    # --------------------------------------------------------
    # Find maximum number of columns
    # --------------------------------------------------------

    maximum_columns = len(
        headers
    )

    for row in rows:

        maximum_columns = max(
            maximum_columns,
            len(row)
        )

    # --------------------------------------------------------
    # Add blank headers if required.
    #
    # These remain blank.
    # No Column 1 / Column 2.
    # --------------------------------------------------------

    while len(headers) < maximum_columns:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # HTML TABLE
    # --------------------------------------------------------

    table_html = """
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<style>

* {
    box-sizing: border-box;
}

body {
    margin: 0;
    padding: 0;

    background: transparent;

    font-family:
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        sans-serif;
}

.table-wrapper {

    width: 100%;

    overflow-x: auto;
    overflow-y: auto;

    max-height: 650px;

    border:
        1px solid #30333d;

    border-radius:
        8px;
}

table {

    border-collapse:
        collapse;

    width:
        max-content;

    min-width:
        100%;

    font-size:
        14px;
}

thead th {

    position:
        sticky;

    top:
        0;

    z-index:
        10;

    background:
        #1f2129;

    color:
        #d8dbe2;

    padding:
        11px 14px;

    border:
        1px solid #3a3d47;

    text-align:
        left;

    white-space:
        nowrap;

    font-weight:
        600;
}

tbody td {

    background:
        #0e1117;

    color:
        #f1f1f1;

    padding:
        10px 14px;

    border:
        1px solid #30333d;

    text-align:
        left;

    white-space:
        nowrap;
}

tbody tr:hover td {

    background:
        #181b23;
}

</style>

</head>

<body>

<div class="table-wrapper">

<table>

<thead>

<tr>
"""

    # ========================================================
    # HEADERS
    # ========================================================

    for header in headers:

        table_html += (
            "<th>"
            + html_value(header)
            + "</th>"
        )

    table_html += """
</tr>

</thead>

<tbody>
"""

    # ========================================================
    # ROWS
    # ========================================================

    for row in rows:

        row = list(row)

        if len(row) < maximum_columns:

            row.extend(
                [None] *
                (
                    maximum_columns - len(row)
                )
            )

        table_html += "<tr>"

        for value in row[
            :maximum_columns
        ]:

            table_html += (
                "<td>"
                + html_value(value)
                + "</td>"
            )

        table_html += "</tr>"

    table_html += """
</tbody>

</table>

</div>

</body>

</html>
"""

    # --------------------------------------------------------
    # Dynamic height
    # --------------------------------------------------------

    height = min(
        650,
        110 + (
            len(rows) * 45
        )
    )

    height = max(
        180,
        height
    )

    # --------------------------------------------------------
    # Render actual HTML
    # --------------------------------------------------------

    components.html(
        table_html,
        height=height,
        scrolling=False
    )


# ============================================================
# CREATE DOWNLOADABLE EXCEL
#
# ONLY the matching result table is downloaded.
#
# Important:
# openpyxl allows duplicate and blank header values.
# ============================================================

def create_download_excel(
    headers,
    rows
):

    output = io.BytesIO()

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Search Result"

    headers = list(
        headers
    )

    # --------------------------------------------------------
    # Determine maximum number of columns
    # --------------------------------------------------------

    maximum_columns = len(
        headers
    )

    for row in rows:

        maximum_columns = max(
            maximum_columns,
            len(row)
        )

    # --------------------------------------------------------
    # Make sure header list covers all data columns
    #
    # Extra headers remain blank.
    # --------------------------------------------------------

    while len(headers) < maximum_columns:

        headers.append(
            None
        )

    # ========================================================
    # WRITE HEADERS
    # ========================================================

    for column_number, header in enumerate(
        headers,
        start=1
    ):

        if header is None:

            cell_value = ""

        else:

            try:

                if pd.isna(header):

                    cell_value = ""

                else:

                    cell_value = header

            except Exception:

                cell_value = header

        worksheet.cell(
            row=1,
            column=column_number,
            value=cell_value
        )

    # ========================================================
    # WRITE MATCHING ROWS
    # ========================================================

    for row_number, row in enumerate(
        rows,
        start=2
    ):

        row = list(
            row
        )

        # Ensure all columns exist
        if len(row) < maximum_columns:

            row.extend(
                [None] *
                (
                    maximum_columns - len(row)
                )
            )

        for column_number, value in enumerate(
            row[:maximum_columns],
            start=1
        ):

            if value is None:

                cell_value = None

            else:

                try:

                    if pd.isna(value):

                        cell_value = None

                    else:

                        cell_value = value

                except Exception:

                    cell_value = value

            worksheet.cell(
                row=row_number,
                column=column_number,
                value=cell_value
            )

    # ========================================================
    # FORMAT DOWNLOAD SHEET
    # ========================================================

    # Freeze header row
    worksheet.freeze_panes = "A2"

    # --------------------------------------------------------
    # Column widths
    # --------------------------------------------------------

    for column_cells in worksheet.columns:

        maximum_length = 0

        for cell in column_cells:

            if cell.value is not None:

                length = len(
                    str(cell.value)
                )

                maximum_length = max(
                    maximum_length,
                    length
                )

        width = min(
            max(
                maximum_length + 2,
                10
            ),
            50
        )

        worksheet.column_dimensions[
            column_cells[
                0
            ].column_letter
        ].width = width

    # ========================================================
    # SAVE
    # ========================================================

    workbook.save(
        output
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# SAFE DOWNLOAD FILENAME
# ============================================================

def make_download_filename(
    filename,
    search_value
):

    # Remove extension
    base_name = re.sub(
        r"\.(xlsx|xlsm)$",
        "",
        filename,
        flags=re.IGNORECASE
    )

    # Remove characters not allowed in Windows filenames
    base_name = re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        base_name
    )

    search_part = re.sub(
        r'[\\/:*?"<>|]+',
        "-",
        search_value
    )

    return (
        f"{base_name}_"
        f"{search_part}_"
        f"Search_Result.xlsx"
    )


# ============================================================
# DISPLAY SEARCH RESULTS
# ============================================================

def display_results(
    results,
    search_value
):

    # --------------------------------------------------------
    # Group by original Excel table
    # --------------------------------------------------------

    grouped = defaultdict(list)

    for result in results:

        key = (
            result["filename"],
            result["sheet"],
            result["header"],
            result["header_row"],
            result["bill_column"]
        )

        grouped[key].append(
            result
        )

    # --------------------------------------------------------
    # Display every matching table
    # --------------------------------------------------------

    for key, group in grouped.items():

        (
            filename,
            sheet_name,
            bill_header,
            header_row,
            bill_column
        ) = key

        st.markdown(
            "---"
        )

        # ----------------------------------------------------
        # File name
        # ----------------------------------------------------

        st.markdown(
            f"### 📄 {filename}"
        )

        # ----------------------------------------------------
        # Information
        # ----------------------------------------------------

        st.write(
            f"**Sheet:** {sheet_name}"
        )

        st.write(
            f"**Bill / Invoice Header:** "
            f"`{bill_header}`"
        )

        st.write(
            f"**Excel Header Row:** "
            f"{header_row}  |  "
            f"**Bill / Invoice Column:** "
            f"{bill_column}"
        )

        # ----------------------------------------------------
        # Headers
        # ----------------------------------------------------

        headers = list(
            group[0]["headers"]
        )

        # ----------------------------------------------------
        # Matching rows
        # ----------------------------------------------------

        rows = [
            item["row"]
            for item in group
        ]

        # ----------------------------------------------------
        # Display table
        # ----------------------------------------------------

        display_excel_table(
            headers,
            rows
        )

        # ====================================================
        # DOWNLOAD TABLE
        # ====================================================

        download_data = create_download_excel(
            headers,
            rows
        )

        download_filename = make_download_filename(
            filename,
            search_value
        )

        # ----------------------------------------------------
        # Download button
        # ----------------------------------------------------

        st.download_button(

            label="⬇️ Download Table",

            data=download_data,

            file_name=download_filename,

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            key=(
                "download_"
                + str(
                    abs(
                        hash(
                            (
                                filename,
                                sheet_name,
                                bill_header,
                                header_row,
                                bill_column,
                                search_value,
                                len(rows)
                            )
                        )
                    )
                )
            ),

            use_container_width=False
        )


# ============================================================
# FILE UPLOADER
# ============================================================

uploaded_files = st.file_uploader(
    "📂 Upload Excel files",

    type=[
        "xlsx",
        "xlsm"
    ],

    accept_multiple_files=True
)


# ============================================================
# MAIN APPLICATION
# ============================================================

if uploaded_files:

    # --------------------------------------------------------
    # Upload status
    # --------------------------------------------------------

    st.success(
        f"✅ {len(uploaded_files)} "
        f"Excel file(s) uploaded"
    )

    # --------------------------------------------------------
    # Store files in memory
    # --------------------------------------------------------

    file_data = tuple(
        (
            file.name,
            file.getvalue()
        )
        for file in uploaded_files
    )

    # --------------------------------------------------------
    # Read and build index
    # --------------------------------------------------------

    with st.spinner(
        "📖 Reading Excel files and building search index..."
    ):

        indexed_files = build_search_index(
            file_data
        )

    # --------------------------------------------------------
    # Show errors
    # --------------------------------------------------------

    for item in indexed_files:

        if item["error"]:

            st.warning(
                f"⚠️ {item['filename']}: "
                f"{item['error']}"
            )

    # --------------------------------------------------------
    # Count detected tables
    # --------------------------------------------------------

    total_tables = sum(
        len(
            item["tables"]
        )
        for item in indexed_files
    )

    # --------------------------------------------------------
    # Detected Bill/Invoice headers
    # --------------------------------------------------------

    if total_tables > 0:

        st.success(
            f"✅ {total_tables} "
            f"Bill / Invoice table(s) detected"
        )

        with st.expander(
            "📋 View detected Bill / Invoice headers"
        ):

            show_detected_headers(
                indexed_files
            )

    else:

        st.warning(
            "⚠️ No Bill / Invoice headers detected."
        )

    # ========================================================
    # SEARCH
    # ========================================================

    st.markdown(
        "### 🔎 Search Bill / Invoice Number"
    )

    search_value = st.text_input(
        "Bill / Invoice Number",

        placeholder="Example: GST/055/18-19",

        label_visibility="collapsed"
    )

    search_button = st.button(
        "🔎 SEARCH",

        type="primary",

        use_container_width=True
    )

    # ========================================================
    # SEARCH ACTION
    # ========================================================

    if search_button:

        # ----------------------------------------------------
        # Empty search
        # ----------------------------------------------------

        if not search_value.strip():

            st.warning(
                "⚠️ Please enter a "
                "Bill / Invoice number."
            )

        # ----------------------------------------------------
        # No detected tables
        # ----------------------------------------------------

        elif total_tables == 0:

            st.error(
                "❌ No Bill / Invoice columns found."
            )

        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        else:

            with st.spinner(
                "🔍 Searching..."
            ):

                results = search_excel(
                    indexed_files,
                    search_value
                )

            # ------------------------------------------------
            # MATCH FOUND
            # ------------------------------------------------

            if results:

                st.success(
                    f"✅ Found "
                    f"{len(results)} "
                    f"exact match(es)"
                )

                display_results(
                    results,
                    search_value
                )

            # ------------------------------------------------
            # NO MATCH
            # ------------------------------------------------

            else:

                st.error(
                    f'❌ No exact match found for '
                    f'"{search_value}"'
                )

                st.caption(
                    "Only detected Bill / Invoice "
                    "columns were searched. "
                    "Matching is exact."
                )

else:

    st.info(
        "👆 Upload one or more Excel files to begin."
    )
