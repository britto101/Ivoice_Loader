import streamlit as st
import io
import re
import html
import unicodedata

from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit.components.v1 as components
from openpyxl import load_workbook, Workbook


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Bill / Invoice Search",
    page_icon="🔎",
    layout="wide"
)


# ============================================================
# TITLE
# ============================================================

st.title("🔎 Bill / Invoice Search")

st.caption(
    "Search exact Bill / Invoice numbers across all uploaded Excel files."
)


# ============================================================
# BASIC HELPERS
# ============================================================

def is_blank(value):

    if value is None:
        return True

    try:
        if value != value:
            return True
    except Exception:
        pass

    return str(value).strip() == ""


def normalize_value(value):

    """
    Converts Excel values into a reliable search string.

    Examples:

        Excel number 74       -> "74"
        Excel 74.0            -> "74"
        Excel "74"            -> "74"
        Excel " 74 "          -> "74"

    Original Excel value is NEVER modified for display.
    """

    if is_blank(value):
        return ""

    # Integer
    if isinstance(value, int):
        return str(value)

    # Float
    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

        return str(value).strip()

    # String
    return str(value).strip()


# ============================================================
# HEADER NORMALIZATION
# ============================================================

def normalize_header(value):

    if is_blank(value):
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value)
    )

    text = text.strip().lower()

    text = re.sub(
        r"[_\-/\\.:]+",
        " ",
        text
    )

    text = re.sub(
        r"[^\w\s]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# BILL / INVOICE HEADER DETECTION
# ============================================================

def is_bill_invoice_header(value):

    header = normalize_header(value)

    if not header:
        return False

    exact_names = {

        "bill no",
        "bill number",
        "bill num",

        "bil no",
        "bil number",
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
        "invoice reference number",

        "inv ref no",
        "inv ref number"
    }

    if header in exact_names:
        return True

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

    return bool(
        words.intersection(bill_words)
        and
        words.intersection(number_words)
    )


# ============================================================
# FIND ALL BILL / INVOICE HEADERS
# ============================================================

def find_bill_headers(rows):

    headers = []

    for row_index, row in enumerate(rows):

        for column_index, value in enumerate(row):

            if is_bill_invoice_header(value):

                headers.append({

                    "row":
                        row_index,

                    "column":
                        column_index,

                    "value":
                        value
                })

    return headers


# ============================================================
# FIND TABLE WIDTH
#
# IMPORTANT:
#
# Your Excel file can contain multiple tables side-by-side.
#
# Example:
#
# A:H  -> first table
# I    -> blank separator
# J:P  -> second table
#
# We determine the table from the header row.
# ============================================================

def find_table_range(
    rows,
    header_row,
    bill_column
):

    if header_row >= len(rows):

        return (
            bill_column,
            bill_column + 1
        )

    header = list(
        rows[
            header_row
        ]
    )

    total_columns = len(
        header
    )

    # --------------------------------------------------------
    # LEFT BOUNDARY
    # --------------------------------------------------------

    start = bill_column

    col = bill_column - 1

    while col >= 0:

        if is_blank(
            header[col]
        ):
            break

        start = col

        col -= 1

    # --------------------------------------------------------
    # RIGHT BOUNDARY
    # --------------------------------------------------------

    end = bill_column

    col = bill_column + 1

    while col < total_columns:

        if is_blank(
            header[col]
        ):
            break

        end = col

        col += 1

    # --------------------------------------------------------
    # Include blank columns INSIDE the detected table.
    #
    # Example:
    #
    # S NO | ITEM | INVOICE NO | SHOP ID | [blank] | GST
    #
    # If there are headers on both sides, preserve the blank.
    # --------------------------------------------------------

    # Search left for first nonblank header
    left_nonblank = bill_column

    while left_nonblank >= 0:

        if not is_blank(
            header[left_nonblank]
        ):
            break

        left_nonblank -= 1

    # Search right for first nonblank header
    right_nonblank = bill_column

    while right_nonblank < total_columns:

        if not is_blank(
            header[right_nonblank]
        ):
            break

        right_nonblank += 1

    if (
        left_nonblank >= 0
        and
        right_nonblank < total_columns
    ):

        start = left_nonblank
        end = right_nonblank

    # --------------------------------------------------------
    # Return Python slice:
    #
    # start inclusive
    # end exclusive
    # --------------------------------------------------------

    return (
        start,
        end + 1
    )


# ============================================================
# FIND NEXT SAME INVOICE HEADER
# ============================================================

def find_section_end(
    rows,
    header_row,
    bill_column,
    all_headers
):

    end_row = len(
        rows
    )

    future_headers = []

    for header in all_headers:

        if header["column"] != bill_column:
            continue

        if header["row"] <= header_row:
            continue

        future_headers.append(
            header["row"]
        )

    if future_headers:

        end_row = min(
            future_headers
        )

    return end_row


# ============================================================
# BUILD TABLE INDEX
# ============================================================

def build_table_index(
    rows,
    filename,
    sheet_name,
    header_info,
    all_headers
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
    # Find horizontal table range
    # --------------------------------------------------------

    start_column, end_column = find_table_range(
        rows,
        header_row,
        bill_column
    )

    # --------------------------------------------------------
    # Find vertical section end
    # --------------------------------------------------------

    end_row = find_section_end(
        rows,
        header_row,
        bill_column,
        all_headers
    )

    # --------------------------------------------------------
    # Actual Excel header row
    # --------------------------------------------------------

    header_row_values = list(
        rows[
            header_row
        ]
    )

    headers = header_row_values[
        start_column:end_column
    ]

    table_width = (
        end_column -
        start_column
    )

    while len(headers) < table_width:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # Relative Bill/Invoice column
    # --------------------------------------------------------

    relative_bill_column = (
        bill_column -
        start_column
    )

    # --------------------------------------------------------
    # SEARCH INDEX
    # --------------------------------------------------------

    search_index = defaultdict(list)

    # --------------------------------------------------------
    # IMPORTANT:
    #
    # Search EVERY ROW in this section.
    #
    # Blank rows are NOT used to stop searching.
    # --------------------------------------------------------

    for row_index in range(
        header_row + 1,
        end_row
    ):

        worksheet_row = list(
            rows[
                row_index
            ]
        )

        # ----------------------------------------------------
        # Make row long enough
        # ----------------------------------------------------

        if len(worksheet_row) < end_column:

            worksheet_row.extend(
                [None] *
                (
                    end_column -
                    len(worksheet_row)
                )
            )

        # ----------------------------------------------------
        # Take ONLY this table
        # ----------------------------------------------------

        table_row = worksheet_row[
            start_column:end_column
        ]

        # ----------------------------------------------------
        # Safety
        # ----------------------------------------------------

        if (
            relative_bill_column < 0
            or
            relative_bill_column >= len(table_row)
        ):
            continue

        # ----------------------------------------------------
        # BILL / INVOICE VALUE
        # ----------------------------------------------------

        invoice_value = table_row[
            relative_bill_column
        ]

        normalized = normalize_value(
            invoice_value
        )

        if normalized == "":
            continue

        # ----------------------------------------------------
        # INDEX IT
        # ----------------------------------------------------

        search_index[
            normalized
        ].append({

            "excel_row":
                row_index + 1,

            "row":
                table_row
        })

    # --------------------------------------------------------
    # Return
    # --------------------------------------------------------

    return {

        "filename":
            filename,

        "sheet":
            sheet_name,

        "header":
            bill_header,

        "header_row":
            header_row + 1,

        "bill_column":
            bill_column + 1,

        "start_column":
            start_column + 1,

        "end_column":
            end_column,

        "headers":
            headers,

        "search_index":
            dict(search_index)
    }


# ============================================================
# READ EXCEL FILE
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

        workbook = load_workbook(
            filename=io.BytesIO(
                file_bytes
            ),
            read_only=True,
            data_only=True
        )

        # ----------------------------------------------------
        # Process every worksheet
        # ----------------------------------------------------

        for worksheet in workbook.worksheets:

            rows = []

            # ------------------------------------------------
            # Read complete worksheet
            # ------------------------------------------------

            for row in worksheet.iter_rows(
                values_only=True
            ):

                rows.append(
                    list(row)
                )

            if not rows:
                continue

            # ------------------------------------------------
            # Detect Bill / Invoice headers
            # ------------------------------------------------

            detected_headers = find_bill_headers(
                rows
            )

            if not detected_headers:
                continue

            # ------------------------------------------------
            # Build separate table for every header
            # ------------------------------------------------

            for header_info in detected_headers:

                table = build_table_index(
                    rows,
                    filename,
                    worksheet.title,
                    header_info,
                    detected_headers
                )

                result[
                    "tables"
                ].append(
                    table
                )

        workbook.close()

    except Exception as error:

        result["error"] = str(
            error
        )

    return result


# ============================================================
# BUILD COMPLETE INDEX
# ============================================================

@st.cache_data(
    show_spinner=False,
    max_entries=5
)
def build_search_index(
    file_data
):

    results = []

    worker_count = min(
        4,
        max(
            1,
            len(file_data)
        )
    )

    with ThreadPoolExecutor(
        max_workers=worker_count
    ) as executor:

        futures = {}

        for filename, file_bytes in file_data:

            future = executor.submit(
                read_excel_file,
                filename,
                file_bytes
            )

            futures[
                future
            ] = filename

        for future in as_completed(
            futures
        ):

            results.append(
                future.result()
            )

    # --------------------------------------------------------
    # Keep alphabetical file order
    # --------------------------------------------------------

    results.sort(
        key=lambda x:
            x["filename"].lower()
    )

    return results


# ============================================================
# SEARCH
# ============================================================

def search_excel(
    indexed_files,
    search_value
):

    search_value = normalize_value(
        search_value
    )

    results = []

    seen = set()

    # --------------------------------------------------------
    # Search every indexed table
    # --------------------------------------------------------

    for file_data in indexed_files:

        for table in file_data[
            "tables"
        ]:

            matches = table[
                "search_index"
            ].get(
                search_value,
                []
            )

            for match in matches:

                excel_row = match[
                    "excel_row"
                ]

                unique_key = (

                    table["filename"],

                    table["sheet"],

                    table["header_row"],

                    table["bill_column"],

                    excel_row
                )

                if unique_key in seen:
                    continue

                seen.add(
                    unique_key
                )

                results.append({

                    "filename":
                        table["filename"],

                    "sheet":
                        table["sheet"],

                    "header":
                        table["header"],

                    "header_row":
                        table["header_row"],

                    "bill_column":
                        table["bill_column"],

                    "start_column":
                        table["start_column"],

                    "end_column":
                        table["end_column"],

                    "excel_row":
                        excel_row,

                    "headers":
                        table["headers"],

                    "row":
                        match["row"]
                })

    return results


# ============================================================
# HTML ESCAPE
# ============================================================

def safe_html(value):

    if is_blank(value):
        return ""

    return html.escape(
        str(value)
    )


# ============================================================
# DISPLAY TABLE
#
# We deliberately do NOT use st.dataframe().
#
# Reason:
#
# Duplicate / blank Excel headers can cause pyarrow errors.
# ============================================================

def display_excel_table(
    headers,
    rows,
    height=650
):

    headers = list(
        headers
    )

    maximum_columns = len(
        headers
    )

    for row in rows:

        maximum_columns = max(
            maximum_columns,
            len(row)
        )

    # --------------------------------------------------------
    # Complete header list
    # --------------------------------------------------------

    while len(headers) < maximum_columns:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # HTML
    # --------------------------------------------------------

    content = """
<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<style>

html,
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

.wrapper {

    width: 100%;

    max-height: 650px;

    overflow-x: auto;

    overflow-y: auto;

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

    table-layout:
        auto;

    font-size:
        14px;
}

th {

    position:
        sticky;

    top:
        0;

    z-index:
        50;

    background:
        #1f2129;

    color:
        #e5e7eb;

    border:
        1px solid #3a3d47;

    padding:
        11px 14px;

    text-align:
        left;

    white-space:
        nowrap;

    font-weight:
        600;

    min-width:
        65px;
}

td {

    background:
        #0e1117;

    color:
        #f5f5f5;

    border:
        1px solid #30333d;

    padding:
        10px 14px;

    text-align:
        left;

    white-space:
        nowrap;

    min-width:
        65px;
}

tr:hover td {

    background:
        #181b23;
}

::-webkit-scrollbar {

    width:
        12px;

    height:
        12px;
}

::-webkit-scrollbar-track {

    background:
        #16181f;
}

::-webkit-scrollbar-thumb {

    background:
        #777;

    border-radius:
        8px;
}

::-webkit-scrollbar-thumb:hover {

    background:
        #999;
}

</style>

</head>

<body>

<div class="wrapper">

<table>

<thead>

<tr>
"""

    # --------------------------------------------------------
    # HEADERS
    # --------------------------------------------------------

    for header in headers:

        content += (
            "<th>"
            + safe_html(header)
            + "</th>"
        )

    content += """

</tr>

</thead>

<tbody>
"""

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    for row in rows:

        row = list(
            row
        )

        if len(row) < maximum_columns:

            row.extend(
                [None] *
                (
                    maximum_columns -
                    len(row)
                )
            )

        content += "<tr>"

        for value in row[
            :maximum_columns
        ]:

            content += (
                "<td>"
                + safe_html(value)
                + "</td>"
            )

        content += "</tr>"

    content += """

</tbody>

</table>

</div>

</body>

</html>
"""

    calculated_height = min(
        height,
        max(
            220,
            100 + len(rows) * 43
        )
    )

    components.html(
        content,
        height=calculated_height,
        scrolling=False
    )


# ============================================================
# DETECTED HEADER TABLE
# ============================================================

def display_detected_headers(
    indexed_files
):

    headers = [
        "File",
        "Sheet",
        "Header",
        "Excel Header Row",
        "Bill / Invoice Column",
        "Table Start",
        "Table End"
    ]

    rows = []

    for file_data in indexed_files:

        for table in file_data[
            "tables"
        ]:

            rows.append([

                table["filename"],

                table["sheet"],

                table["header"],

                table["header_row"],

                table["bill_column"],

                table["start_column"],

                table["end_column"]
            ])

    if rows:

        display_excel_table(
            headers,
            rows,
            height=450
        )


# ============================================================
# CREATE DOWNLOAD EXCEL
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

    maximum_columns = len(
        headers
    )

    for row in rows:

        maximum_columns = max(
            maximum_columns,
            len(row)
        )

    while len(headers) < maximum_columns:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # Headers
    # --------------------------------------------------------

    for column_number, header in enumerate(
        headers,
        start=1
    ):

        worksheet.cell(
            row=1,
            column=column_number,
            value=(
                ""
                if is_blank(header)
                else header
            )
        )

    # --------------------------------------------------------
    # Data
    # --------------------------------------------------------

    for row_number, row in enumerate(
        rows,
        start=2
    ):

        row = list(
            row
        )

        if len(row) < maximum_columns:

            row.extend(
                [None] *
                (
                    maximum_columns -
                    len(row)
                )
            )

        for column_number, value in enumerate(
            row[
                :maximum_columns
            ],
            start=1
        ):

            worksheet.cell(
                row=row_number,
                column=column_number,
                value=(
                    None
                    if is_blank(value)
                    else value
                )
            )

    # --------------------------------------------------------
    # Freeze header
    # --------------------------------------------------------

    worksheet.freeze_panes = "A2"

    # --------------------------------------------------------
    # Width
    # --------------------------------------------------------

    for column_cells in worksheet.columns:

        maximum_length = 0

        for cell in column_cells:

            if cell.value is not None:

                maximum_length = max(
                    maximum_length,
                    len(
                        str(cell.value)
                    )
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

    # --------------------------------------------------------
    # Save
    # --------------------------------------------------------

    workbook.save(
        output
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# DOWNLOAD FILE NAME
# ============================================================

def download_filename(
    original_filename,
    search_value
):

    base = re.sub(
        r"\.(xlsx|xlsm)$",
        "",
        original_filename,
        flags=re.IGNORECASE
    )

    base = re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        base
    )

    search_part = re.sub(
        r'[\\/:*?"<>|]+',
        "-",
        search_value
    )

    return (
        f"{base}_"
        f"{search_part}_"
        f"Search_Result.xlsx"
    )


# ============================================================
# DISPLAY RESULTS
# ============================================================

def display_results(
    results,
    search_value
):

    # --------------------------------------------------------
    # Group by actual Excel table
    # --------------------------------------------------------

    grouped = defaultdict(list)

    for result in results:

        key = (

            result["filename"],

            result["sheet"],

            result["header"],

            result["header_row"],

            result["bill_column"],

            result["start_column"],

            result["end_column"]
        )

        grouped[
            key
        ].append(
            result
        )

    # --------------------------------------------------------
    # Display groups
    # --------------------------------------------------------

    for group_number, (
        key,
        group
    ) in enumerate(
        grouped.items()
    ):

        (
            filename,
            sheet_name,
            bill_header,
            header_row,
            bill_column,
            start_column,
            end_column
        ) = key

        st.markdown(
            "---"
        )

        # ----------------------------------------------------
        # File
        # ----------------------------------------------------

        st.markdown(
            f"### 📄 {filename}"
        )

        # ----------------------------------------------------
        # Sheet
        # ----------------------------------------------------

        st.write(
            f"**Sheet:** {sheet_name}"
        )

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        st.write(
            f"**Bill / Invoice Header:** "
            f"`{bill_header}`"
        )

        # ----------------------------------------------------
        # Location
        # ----------------------------------------------------

        st.write(
            f"**Excel Header Row:** "
            f"{header_row}"
            f"  |  "
            f"**Bill / Invoice Column:** "
            f"{bill_column}"
        )

        # ----------------------------------------------------
        # Table range
        # ----------------------------------------------------

        st.caption(
            f"Table columns: "
            f"{start_column} → {end_column}"
        )

        # ----------------------------------------------------
        # Rows
        # ----------------------------------------------------

        rows = [
            result["row"]
            for result in group
        ]

        headers = list(
            group[0]["headers"]
        )

        st.caption(
            f"Showing {len(rows)} "
            f"complete matching Excel row(s)."
        )

        # ----------------------------------------------------
        # DISPLAY
        # ----------------------------------------------------

        display_excel_table(
            headers,
            rows,
            height=650
        )

        # ----------------------------------------------------
        # DOWNLOAD
        # ----------------------------------------------------

        download_data = create_download_excel(
            headers,
            rows
        )

        file_name = download_filename(
            filename,
            search_value
        )

        st.download_button(

            "⬇️ Download Table",

            data=download_data,

            file_name=file_name,

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
                                header_row,
                                bill_column,
                                start_column,
                                end_column,
                                search_value,
                                group_number
                            )
                        )
                    )
                )
            )
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
# MAIN
# ============================================================

if uploaded_files:

    # ========================================================
    # UPLOAD STATUS
    # ========================================================

    st.success(
        f"✅ {len(uploaded_files)} "
        f"Excel file(s) uploaded"
    )

    # ========================================================
    # GET FILE BYTES
    # ========================================================

    file_data = tuple(

        (
            file.name,
            file.getvalue()
        )

        for file in uploaded_files
    )

    # ========================================================
    # READ / INDEX
    # ========================================================

    with st.spinner(
        "📖 Reading Excel files..."
    ):

        indexed_files = build_search_index(
            file_data
        )

    # ========================================================
    # ERRORS
    # ========================================================

    for file_data_result in indexed_files:

        if file_data_result["error"]:

            st.error(
                f"❌ "
                f"{file_data_result['filename']}: "
                f"{file_data_result['error']}"
            )

    # ========================================================
    # TABLE COUNT
    # ========================================================

    total_tables = sum(

        len(
            item["tables"]
        )

        for item in indexed_files
    )

    if total_tables:

        st.success(
            f"✅ {total_tables} "
            f"Bill / Invoice table(s) detected"
        )

        # ----------------------------------------------------
        # HEADER LIST
        # ----------------------------------------------------

        with st.expander(
            "📋 View detected Bill / Invoice headers"
        ):

            display_detected_headers(
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
        "## 🔎 Search Bill / Invoice Number"
    )

    search_value = st.text_input(

        "Bill / Invoice Number",

        placeholder=(
            "Example: "
            "74 or CRSCL/19-20/1086"
        ),

        label_visibility="collapsed"
    )

    search_button = st.button(

        "🔎 SEARCH",

        type="primary",

        use_container_width=True
    )

    # ========================================================
    # SEARCH BUTTON
    # ========================================================

    if search_button:

        if not search_value.strip():

            st.warning(
                "⚠️ Please enter a "
                "Bill / Invoice number."
            )

        elif total_tables == 0:

            st.error(
                "❌ No Bill / Invoice "
                "columns detected."
            )

        else:

            with st.spinner(
                "🔍 Searching..."
            ):

                results = search_excel(
                    indexed_files,
                    search_value
                )

            # ------------------------------------------------
            # FOUND
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
            # NOT FOUND
            # ------------------------------------------------

            else:

                st.error(
                    f'❌ No exact match found for '
                    f'"{search_value}"'
                )

                st.caption(
                    "The search checks only detected "
                    "Bill / Invoice columns. "
                    "Numbers such as 74 are matched "
                    "correctly whether Excel stores "
                    "them as numbers or text."
                )

else:

    st.info(
        "👆 Upload one or more Excel files to begin."
    )
