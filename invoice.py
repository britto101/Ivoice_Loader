Pasted text(20260902-155201).txt
Document
not printing when i put this bill no because inbetween 2 column empty make or add this alone dont change remaining things in ocde

Yes. The problem is in find_table_boundaries(). Your sheet has blank columns between BIL NO and ITEMS, so the current code stops at the first blank column. That is why the matching rows are not printed correctly.

Only change this one function

Replace your existing find_table_boundaries() function with this:

# ============================================================
# DETERMINE TABLE BOUNDARIES
#
# IMPORTANT:
# Excel tables may contain blank columns between headers.
# Example:
#
# S NO | BIL NO | BLANK | BLANK | ITEMS | SHOP ID
#
# We must keep those blank columns as part of the table.
# ============================================================

def find_table_boundaries(
    rows,
    header_row,
    invoice_column
):

    if header_row >= len(rows):
        return (
            invoice_column,
            invoice_column + 1
        )

    header = list(rows[header_row])
    column_count = len(header)

    # --------------------------------------------------------
    # Find first meaningful header on the LEFT
    # --------------------------------------------------------

    start = invoice_column

    for column in range(invoice_column - 1, -1, -1):

        if not is_blank(header[column]):
            start = column

        # Keep going through blank columns.
        # Stop only after finding the first meaningful
        # header to the left.
        else:
            continue

    # --------------------------------------------------------
    # Find LAST meaningful header on the RIGHT
    #
    # IMPORTANT:
    # Do NOT stop at blank columns.
    # Blank columns between BIL NO / ITEMS / SHOP ID
    # remain part of the table.
    # --------------------------------------------------------

    end = invoice_column

    for column in range(invoice_column + 1, column_count):

        if not is_blank(header[column]):
            end = column

    # --------------------------------------------------------
    # Return Python slice
    # --------------------------------------------------------

    return (
        start,
        end + 1
    )
What this fixes

Your Excel structure:

S NO | BIL NO |       |       | ITEMS | SHOP ID
 1   | SAL/...|       |       |SYSTEMS| 15221
 2   | SAL/...|       |       |SYSTEMS| 15284

will now be treated as one complete table:

S NO | BIL NO | blank | blank | ITEMS   | SHOP ID
  1  | SAL/...|       |       | SYSTEMS | 15221
  2  | SAL/...|       |       | SYSTEMS | 15284

So searching:

SAL/083/19-20

will print both matching rows, including the blank columns and ITEMS / SHOP ID.

Nothing else in your code needs to be changed. Keep the rest exactly as it is.

give full corrected code

Absolutely. Below is the full corrected code. I have kept the existing functionality and styling the same, with the table-boundary logic changed so that blank columns between BIL NO, ITEMS, SHOP ID, etc. are included.

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
# PAGE CONFIGURATION
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
    "Search Bill / Invoice numbers across all uploaded Excel files."
)


# ============================================================
# BASIC FUNCTIONS
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

    if is_blank(value):
        return ""

    # Excel integer
    if isinstance(value, int):
        return str(value)

    # Excel float
    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

        return str(value).strip()

    # Everything else
    return str(value).strip()


# ============================================================
# NORMALIZE HEADER
# ============================================================

def normalize_header(value):

    if is_blank(value):
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value)
    )

    text = text.strip().lower()

    # Remove common separators
    text = re.sub(
        r"[_\-/\\.:]+",
        " ",
        text
    )

    # Remove punctuation
    text = re.sub(
        r"[^\w\s]",
        " ",
        text
    )

    # Multiple spaces
    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# BILL / INVOICE HEADER DETECTION
#
# Supports:
#
# BILL
# BILL NO
# BILL NUMBER
# BIL NO
# INVOICE
# INVOICE NO
# INVOICE NUMBER
# INOVICE NO       <-- typo
# INOVICE NUMBER   <-- typo
# INV
# INV NO
# INV NUMBER
# ============================================================

def is_bill_invoice_header(value):

    header = normalize_header(value)

    if not header:
        return False

    # Exact supported names
    exact_names = {

        # BILL
        "bill",
        "bill no",
        "bill number",
        "bill num",

        "bil",
        "bil no",
        "bil number",
        "bil num",

        # INVOICE
        "invoice",
        "invoice no",
        "invoice number",
        "invoice num",

        # Common Excel typo
        "inovice",
        "inovice no",
        "inovice number",
        "inovice num",

        # INV
        "inv",
        "inv no",
        "inv number",
        "inv num",

        # Reference formats
        "bill ref",
        "bill ref no",
        "bill ref number",

        "invoice ref",
        "invoice ref no",
        "invoice ref number",

        "inovice ref",
        "inovice ref no",
        "inovice ref number"
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
        "inovice",
        "inv"
    }

    number_words = {
        "no",
        "number",
        "num"
    }

    # Example:
    #
    # BILL NO
    # INVOICE NO
    # INOVICE NO
    # INV NUMBER
    #

    if (
        words.intersection(bill_words)
        and
        words.intersection(number_words)
    ):
        return True

    return False


# ============================================================
# FIND ALL BILL / INVOICE HEADERS
# ============================================================

def find_bill_invoice_headers(rows):

    headers = []

    for row_index, row in enumerate(rows):

        for column_index, value in enumerate(row):

            if is_bill_invoice_header(value):

                headers.append({
                    "row": row_index,
                    "column": column_index,
                    "value": value
                })

    return headers


# ============================================================
# DETERMINE TABLE BOUNDARIES
#
# IMPORTANT FIX:
#
# Excel tables can have blank columns between headers.
#
# Example:
#
# S NO | BIL NO |       |       | ITEMS | SHOP ID
#
# The blank columns must remain part of the table.
#
# ============================================================

def find_table_boundaries(
    rows,
    header_row,
    invoice_column
):

    if header_row >= len(rows):

        return (
            invoice_column,
            invoice_column + 1
        )

    header = list(
        rows[header_row]
    )

    column_count = len(header)

    # --------------------------------------------------------
    # Find first meaningful header on the LEFT
    # --------------------------------------------------------

    start = invoice_column

    for column in range(
        invoice_column - 1,
        -1,
        -1
    ):

        if not is_blank(header[column]):

            start = column

        # IMPORTANT:
        # Do not stop at blank columns.
        #
        # Blank columns may belong to the same table.
        #
        else:
            continue

    # --------------------------------------------------------
    # Find last meaningful header on the RIGHT
    # --------------------------------------------------------

    end = invoice_column

    for column in range(
        invoice_column + 1,
        column_count
    ):

        if not is_blank(header[column]):

            end = column

        # IMPORTANT:
        # Continue through blank columns.
        #
        # This allows:
        #
        # BIL NO | blank | blank | ITEMS | SHOP ID
        #
        else:
            continue

    # --------------------------------------------------------
    # Return Python slice
    # --------------------------------------------------------

    return (
        start,
        end + 1
    )


# ============================================================
# FIND END OF THIS TABLE
#
# A table ends when:
#
# 1. Another header row for the SAME invoice column appears
# 2. A new table with a different invoice column/header appears
#    after sufficient separation
#
# We do NOT stop simply because there is one blank row.
# ============================================================

def find_table_end(
    rows,
    header_row,
    invoice_column,
    all_headers
):

    total_rows = len(rows)

    possible_next_headers = []

    for header in all_headers:

        if header["row"] <= header_row:
            continue

        possible_next_headers.append(
            header
        )

    # --------------------------------------------------------
    # Look for the next invoice/bill header.
    #
    # This is the safest boundary because the workbook can
    # contain multiple tables vertically.
    # --------------------------------------------------------

    for header in sorted(
        possible_next_headers,
        key=lambda x: x["row"]
    ):

        if header["row"] > header_row:

            return header["row"]

    return total_rows


# ============================================================
# BUILD ONE TABLE
# ============================================================

def build_table(
    rows,
    filename,
    sheet_name,
    header_info,
    all_headers
):

    header_row = header_info[
        "row"
    ]

    invoice_column = header_info[
        "column"
    ]

    invoice_header = header_info[
        "value"
    ]

    # --------------------------------------------------------
    # Horizontal range
    # --------------------------------------------------------

    start_column, end_column = find_table_boundaries(
        rows,
        header_row,
        invoice_column
    )

    # --------------------------------------------------------
    # Vertical range
    # --------------------------------------------------------

    end_row = find_table_end(
        rows,
        header_row,
        invoice_column,
        all_headers
    )

    # --------------------------------------------------------
    # Get EXACT Excel header names
    # --------------------------------------------------------

    source_header_row = list(
        rows[header_row]
    )

    headers = source_header_row[
        start_column:end_column
    ]

    expected_width = (
        end_column -
        start_column
    )

    while len(headers) < expected_width:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # Invoice column inside this table
    # --------------------------------------------------------

    relative_invoice_column = (
        invoice_column -
        start_column
    )

    # --------------------------------------------------------
    # SEARCH INDEX
    # --------------------------------------------------------

    search_index = defaultdict(list)

    # --------------------------------------------------------
    # READ ALL ROWS BELONGING TO THIS TABLE
    # --------------------------------------------------------

    for row_number in range(
        header_row + 1,
        end_row
    ):

        source_row = list(
            rows[row_number]
        )

        # Make sure row has enough columns
        if len(source_row) < end_column:

            source_row.extend(
                [None] *
                (
                    end_column -
                    len(source_row)
                )
            )

        # ----------------------------------------------------
        # ONLY THIS TABLE'S COLUMNS
        # ----------------------------------------------------

        table_row = source_row[
            start_column:end_column
        ]

        # ----------------------------------------------------
        # Safety
        # ----------------------------------------------------

        if (
            relative_invoice_column < 0
            or
            relative_invoice_column >= len(table_row)
        ):
            continue

        # ----------------------------------------------------
        # Invoice value
        # ----------------------------------------------------

        invoice_value = table_row[
            relative_invoice_column
        ]

        normalized = normalize_value(
            invoice_value
        )

        if normalized == "":
            continue

        # ----------------------------------------------------
        # Store complete table row
        # ----------------------------------------------------

        search_index[
            normalized
        ].append({

            "excel_row":
                row_number + 1,

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
            invoice_header,

        "header_row":
            header_row + 1,

        "invoice_column":
            invoice_column + 1,

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

        workbook = load_workbook(
            filename=io.BytesIO(
                file_bytes
            ),
            read_only=True,
            data_only=True
        )

        # ----------------------------------------------------
        # Every sheet
        # ----------------------------------------------------

        for worksheet in workbook.worksheets:

            rows = []

            # ------------------------------------------------
            # Read complete sheet
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
            # Detect every Bill / Invoice header
            # ------------------------------------------------

            headers = find_bill_invoice_headers(
                rows
            )

            if not headers:
                continue

            # ------------------------------------------------
            # Every header = separate table
            # ------------------------------------------------

            for header_info in headers:

                table = build_table(
                    rows,
                    filename,
                    worksheet.title,
                    header_info,
                    headers
                )

                if table:

                    result[
                        "tables"
                    ].append(
                        table
                    )

        workbook.close()

    except Exception as error:

        result[
            "error"
        ] = str(error)

    return result


# ============================================================
# BUILD SEARCH INDEX
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

                unique_key = (

                    table["filename"],

                    table["sheet"],

                    table["header_row"],

                    table["invoice_column"],

                    match["excel_row"]
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

                    "invoice_column":
                        table["invoice_column"],

                    "start_column":
                        table["start_column"],

                    "end_column":
                        table["end_column"],

                    "excel_row":
                        match["excel_row"],

                    "headers":
                        table["headers"],

                    "row":
                        match["row"]
                })

    return results


# ============================================================
# SAFE HTML
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
# We intentionally DO NOT use st.dataframe()
# because Excel files can contain:
#
# - duplicate headers
# - blank headers
#
# which can cause PyArrow errors.
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
    # Preserve blank Excel headers.
    #
    # DO NOT rename them to:
    #
    # Column 1
    # Column 2
    #
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

        table_html += (
            "<th>"
            +
            safe_html(header)
            +
            "</th>"
        )

    table_html += """

</tr>

</thead>

<tbody>
"""

    # --------------------------------------------------------
    # ROWS
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

        table_html += "<tr>"

        for value in row[
            :maximum_columns
        ]:

            table_html += (
                "<td>"
                +
                safe_html(value)
                +
                "</td>"
            )

        table_html += "</tr>"

    table_html += """

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
        table_html,
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

        "Bill / Invoice Header",

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

                table["invoice_column"],

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
# DOWNLOAD EXCEL
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
    # Header row
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
            row[:maximum_columns],
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
    # Auto width
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

def make_download_filename(
    filename,
    search_value
):

    base = re.sub(
        r"\.(xlsx|xlsm)$",
        "",
        filename,
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
# DISPLAY SEARCH RESULTS
# ============================================================

def display_results(
    results,
    search_value
):

    # --------------------------------------------------------
    # Group matches by ACTUAL TABLE
    # --------------------------------------------------------

    grouped = defaultdict(list)

    for result in results:

        key = (

            result["filename"],

            result["sheet"],

            result["header"],

            result["header_row"],

            result["invoice_column"],

            result["start_column"],

            result["end_column"]
        )

        grouped[
            key
        ].append(
            result
        )

    # --------------------------------------------------------
    # Display each table
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
            invoice_header,
            header_row,
            invoice_column,
            start_column,
            end_column
        ) = key

        st.markdown(
            "---"
        )

        st.markdown(
            f"### 📄 {filename}"
        )

        st.write(
            f"**Sheet:** {sheet_name}"
        )

        st.write(
            f"**Bill / Invoice Header:** "
            f"`{invoice_header}`"
        )

        st.write(
            f"**Excel Header Row:** "
            f"{header_row}"
            f"  |  "
            f"**Bill / Invoice Column:** "
            f"{invoice_column}"
        )

        st.caption(
            f"Table columns: "
            f"{start_column} → {end_column}"
        )

        st.caption(
            f"Showing {len(group)} "
            f"complete matching Excel row(s)."
        )

        # ----------------------------------------------------
        # EXACT HEADERS
        # ----------------------------------------------------

        headers = list(
            group[0]["headers"]
        )

        # ----------------------------------------------------
        # EXACT MATCHING ROWS
        # ----------------------------------------------------

        rows = [

            result["row"]

            for result in group
        ]

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

        download_name = make_download_filename(
            filename,
            search_value
        )

        st.download_button(

            "⬇️ Download Table",

            data=download_data,

            file_name=download_name,

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            key=(
                "download_"
                +
                str(
                    abs(
                        hash(
                            (
                                filename,
                                sheet_name,
                                invoice_header,
                                header_row,
                                invoice_column,
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
# UPLOAD
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
# APPLICATION
# ============================================================

if uploaded_files:

    st.success(
        f"✅ {len(uploaded_files)} "
        f"Excel file(s) uploaded"
    )

    # --------------------------------------------------------
    # Store uploaded files in memory
    # --------------------------------------------------------

    file_data = tuple(

        (
            file.name,
            file.getvalue()
        )

        for file in uploaded_files
    )

    # --------------------------------------------------------
    # Build index
    # --------------------------------------------------------

    with st.spinner(
        "📖 Reading Excel files..."
    ):

        indexed_files = build_search_index(
            file_data
        )

    # --------------------------------------------------------
    # Errors
    # --------------------------------------------------------

    for file_result in indexed_files:

        if file_result["error"]:

            st.error(
                f"❌ "
                f"{file_result['filename']}: "
                f"{file_result['error']}"
            )

    # --------------------------------------------------------
    # Table count
    # --------------------------------------------------------

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
    # SEARCH AREA
    # ========================================================

    st.markdown(
        "## 🔎 Search Bill / Invoice Number"
    )

    search_value = st.text_input(

        "Bill / Invoice Number",

        placeholder=(
            "Example: 74, 162, "
            "CRSCL/19-20/1086, "
            "SMS/0504/2021-22"
        ),

        label_visibility="collapsed"
    )

    search_button = st.button(

        "🔎 SEARCH",

        type="primary",

        use_container_width=True
    )

    # ========================================================
    # SEARCH
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
                    "Search is performed only in "
                    "detected Bill / Invoice columns."
                )

else:

    st.info(
        "👆 Upload one or more Excel files to begin."
    )
The important change

Your original code stopped when it encountered a blank header between columns. The corrected version now keeps those blank columns, so this structure:

S NO | BIL NO | [BLANK] | [BLANK] | ITEMS | SHOP ID

is treated as one table.

So for your example:

SAL/083/19-20

it should return:

1 | SAL/083/19-20 |       |       | SYSTEMS | 15221
2 | SAL/083/19-20 |       |       | SYSTEMS | 15284

and similarly:

SAL/049/19-20

should return the matching rows from that section, including the blank columns. The rest of the search, display, dark table, and Excel download functionality remains unchanged.

here until shop only printing give full modified code

Yes — I can see the exact issue from your screenshot. The result is stopping at SHOP NAME because the table width is being determined from the header row instead of the complete data area. Your actual table continues through:

CAMERA(10) → BNC 20 → ADAPTER10 → DVR 5 → HARDDISK 5

I changed the table detection so it determines the last used column from the entire table's rows, while still preserving blank columns. Everything else remains the same.

Full corrected code
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
# PAGE CONFIGURATION
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
    "Search Bill / Invoice numbers across all uploaded Excel files."
)


# ============================================================
# BASIC FUNCTIONS
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

    if is_blank(value):
        return ""

    # Excel integer
    if isinstance(value, int):
        return str(value)

    # Excel float
    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

        return str(value).strip()

    # Everything else
    return str(value).strip()


# ============================================================
# NORMALIZE HEADER
# ============================================================

def normalize_header(value):

    if is_blank(value):
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value)
    )

    text = text.strip().lower()

    # Remove common separators
    text = re.sub(
        r"[_\-/\\.:]+",
        " ",
        text
    )

    # Remove punctuation
    text = re.sub(
        r"[^\w\s]",
        " ",
        text
    )

    # Multiple spaces
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

        # BILL
        "bill",
        "bill no",
        "bill number",
        "bill num",

        "bil",
        "bil no",
        "bil number",
        "bil num",

        # INVOICE
        "invoice",
        "invoice no",
        "invoice number",
        "invoice num",

        # Common typo
        "inovice",
        "inovice no",
        "inovice number",
        "inovice num",

        # INV
        "inv",
        "inv no",
        "inv number",
        "inv num",

        # Reference formats
        "bill ref",
        "bill ref no",
        "bill ref number",

        "invoice ref",
        "invoice ref no",
        "invoice ref number",

        "inovice ref",
        "inovice ref no",
        "inovice ref number"
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
        "inovice",
        "inv"
    }

    number_words = {
        "no",
        "number",
        "num"
    }

    if (
        words.intersection(bill_words)
        and
        words.intersection(number_words)
    ):
        return True

    return False


# ============================================================
# FIND ALL BILL / INVOICE HEADERS
# ============================================================

def find_bill_invoice_headers(rows):

    headers = []

    for row_index, row in enumerate(rows):

        for column_index, value in enumerate(row):

            if is_bill_invoice_header(value):

                headers.append({
                    "row": row_index,
                    "column": column_index,
                    "value": value
                })

    return headers


# ============================================================
# FIND END OF THIS TABLE
# ============================================================

def find_table_end(
    rows,
    header_row,
    invoice_column,
    all_headers
):

    total_rows = len(rows)

    possible_next_headers = []

    for header in all_headers:

        if header["row"] <= header_row:
            continue

        possible_next_headers.append(
            header
        )

    # --------------------------------------------------------
    # The next Bill / Invoice header marks the next table.
    # --------------------------------------------------------

    for header in sorted(
        possible_next_headers,
        key=lambda x: x["row"]
    ):

        if header["row"] > header_row:

            return header["row"]

    return total_rows


# ============================================================
# DETERMINE TABLE BOUNDARIES
#
# IMPORTANT:
#
# The Excel sheet may contain:
#
# SNO | BILL NUMBER | SHOP ID | SHOP NAME |
# CAMERA | BNC | ADAPTER | DVR | HARDDISK
#
# The data may also contain blank columns.
#
# We therefore determine the table width from BOTH:
#
# 1. Header row
# 2. All data rows belonging to this table
#
# This prevents the table from stopping at SHOP NAME.
# ============================================================

def find_table_boundaries(
    rows,
    header_row,
    invoice_column,
    end_row
):

    if header_row >= len(rows):

        return (
            invoice_column,
            invoice_column + 1
        )

    # --------------------------------------------------------
    # START COLUMN
    # --------------------------------------------------------

    header = list(
        rows[header_row]
    )

    start = invoice_column

    # Find first meaningful header on the left.
    # Blank columns are allowed.

    for column in range(
        invoice_column - 1,
        -1,
        -1
    ):

        if column >= len(header):
            continue

        if not is_blank(header[column]):

            start = column

    # --------------------------------------------------------
    # END COLUMN
    #
    # IMPORTANT FIX
    #
    # Do NOT rely only on the header row.
    #
    # Search every row belonging to this table and find
    # the furthest used column.
    # --------------------------------------------------------

    end = invoice_column

    search_end = min(
        end_row,
        len(rows)
    )

    for row_number in range(
        header_row,
        search_end
    ):

        row = list(
            rows[row_number]
        )

        for column in range(
            len(row) - 1,
            end - 1,
            -1
        ):

            if not is_blank(row[column]):

                end = max(
                    end,
                    column
                )

                break

    # --------------------------------------------------------
    # Return Python slice
    # --------------------------------------------------------

    return (
        start,
        end + 1
    )


# ============================================================
# BUILD ONE TABLE
# ============================================================

def build_table(
    rows,
    filename,
    sheet_name,
    header_info,
    all_headers
):

    header_row = header_info[
        "row"
    ]

    invoice_column = header_info[
        "column"
    ]

    invoice_header = header_info[
        "value"
    ]

    # --------------------------------------------------------
    # FIRST determine vertical table boundary
    # --------------------------------------------------------

    end_row = find_table_end(
        rows,
        header_row,
        invoice_column,
        all_headers
    )

    # --------------------------------------------------------
    # THEN determine complete horizontal table boundary
    #
    # This is the important fix.
    # --------------------------------------------------------

    start_column, end_column = find_table_boundaries(
        rows,
        header_row,
        invoice_column,
        end_row
    )

    # --------------------------------------------------------
    # Get EXACT Excel header names
    # --------------------------------------------------------

    source_header_row = list(
        rows[header_row]
    )

    # Make sure header row contains all required columns
    if len(source_header_row) < end_column:

        source_header_row.extend(
            [None] *
            (
                end_column -
                len(source_header_row)
            )
        )

    headers = source_header_row[
        start_column:end_column
    ]

    expected_width = (
        end_column -
        start_column
    )

    while len(headers) < expected_width:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # Invoice column inside this table
    # --------------------------------------------------------

    relative_invoice_column = (
        invoice_column -
        start_column
    )

    # --------------------------------------------------------
    # SEARCH INDEX
    # --------------------------------------------------------

    search_index = defaultdict(list)

    # --------------------------------------------------------
    # READ ALL ROWS BELONGING TO THIS TABLE
    # --------------------------------------------------------

    for row_number in range(
        header_row + 1,
        end_row
    ):

        source_row = list(
            rows[row_number]
        )

        # ----------------------------------------------------
        # Make sure row has enough columns
        # ----------------------------------------------------

        if len(source_row) < end_column:

            source_row.extend(
                [None] *
                (
                    end_column -
                    len(source_row)
                )
            )

        # ----------------------------------------------------
        # ONLY THIS TABLE'S COLUMNS
        # ----------------------------------------------------

        table_row = source_row[
            start_column:end_column
        ]

        # ----------------------------------------------------
        # Safety
        # ----------------------------------------------------

        if (
            relative_invoice_column < 0
            or
            relative_invoice_column >= len(table_row)
        ):
            continue

        # ----------------------------------------------------
        # Invoice value
        # ----------------------------------------------------

        invoice_value = table_row[
            relative_invoice_column
        ]

        normalized = normalize_value(
            invoice_value
        )

        if normalized == "":
            continue

        # ----------------------------------------------------
        # Store COMPLETE table row
        # ----------------------------------------------------

        search_index[
            normalized
        ].append({

            "excel_row":
                row_number + 1,

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
            invoice_header,

        "header_row":
            header_row + 1,

        "invoice_column":
            invoice_column + 1,

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

        workbook = load_workbook(
            filename=io.BytesIO(
                file_bytes
            ),
            read_only=True,
            data_only=True
        )

        # ----------------------------------------------------
        # Every sheet
        # ----------------------------------------------------

        for worksheet in workbook.worksheets:

            rows = []

            # ------------------------------------------------
            # Read COMPLETE sheet
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
            # Detect every Bill / Invoice header
            # ------------------------------------------------

            headers = find_bill_invoice_headers(
                rows
            )

            if not headers:
                continue

            # ------------------------------------------------
            # Every header = separate table
            # ------------------------------------------------

            for header_info in headers:

                table = build_table(
                    rows,
                    filename,
                    worksheet.title,
                    header_info,
                    headers
                )

                if table:

                    result[
                        "tables"
                    ].append(
                        table
                    )

        workbook.close()

    except Exception as error:

        result[
            "error"
        ] = str(error)

    return result


# ============================================================
# BUILD SEARCH INDEX
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

                unique_key = (

                    table["filename"],

                    table["sheet"],

                    table["header_row"],

                    table["invoice_column"],

                    match["excel_row"]
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

                    "invoice_column":
                        table["invoice_column"],

                    "start_column":
                        table["start_column"],

                    "end_column":
                        table["end_column"],

                    "excel_row":
                        match["excel_row"],

                    "headers":
                        table["headers"],

                    "row":
                        match["row"]
                })

    return results


# ============================================================
# SAFE HTML
# ============================================================

def safe_html(value):

    if is_blank(value):
        return ""

    return html.escape(
        str(value)
    )


# ============================================================
# DISPLAY TABLE
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
    # Preserve blank Excel headers
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

        table_html += (
            "<th>"
            +
            safe_html(header)
            +
            "</th>"
        )

    table_html += """

</tr>

</thead>

<tbody>
"""

    # --------------------------------------------------------
    # ROWS
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

        table_html += "<tr>"

        for value in row[
            :maximum_columns
        ]:

            table_html += (
                "<td>"
                +
                safe_html(value)
                +
                "</td>"
            )

        table_html += "</tr>"

    table_html += """

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
        table_html,
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

        "Bill / Invoice Header",

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

                table["invoice_column"],

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
# DOWNLOAD EXCEL
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
    # Header row
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
            row[:maximum_columns],
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
    # Auto width
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

def make_download_filename(
    filename,
    search_value
):

    base = re.sub(
        r"\.(xlsx|xlsm)$",
        "",
        filename,
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
# DISPLAY SEARCH RESULTS
# ============================================================

def display_results(
    results,
    search_value
):

    # --------------------------------------------------------
    # Group matches by ACTUAL TABLE
    # --------------------------------------------------------

    grouped = defaultdict(list)

    for result in results:

        key = (

            result["filename"],

            result["sheet"],

            result["header"],

            result["header_row"],

            result["invoice_column"],

            result["start_column"],

            result["end_column"]
        )

        grouped[
            key
        ].append(
            result
        )

    # --------------------------------------------------------
    # Display each table
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
            invoice_header,
            header_row,
            invoice_column,
            start_column,
            end_column
        ) = key

        st.markdown(
            "---"
        )

        st.markdown(
            f"### 📄 {filename}"
        )

        st.write(
            f"**Sheet:** {sheet_name}"
        )

        st.write(
            f"**Bill / Invoice Header:** "
            f"`{invoice_header}`"
        )

        st.write(
            f"**Excel Header Row:** "
            f"{header_row}"
            f"  |  "
            f"**Bill / Invoice Column:** "
            f"{invoice_column}"
        )

        st.caption(
            f"Table columns: "
            f"{start_column} → {end_column}"
        )

        st.caption(
            f"Showing {len(group)} "
            f"complete matching Excel row(s)."
        )

        # ----------------------------------------------------
        # EXACT HEADERS
        # ----------------------------------------------------

        headers = list(
            group[0]["headers"]
        )

        # ----------------------------------------------------
        # EXACT MATCHING ROWS
        # ----------------------------------------------------

        rows = [

            result["row"]

            for result in group
        ]

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

        download_name = make_download_filename(
            filename,
            search_value
        )

        st.download_button(

            "⬇️ Download Table",

            data=download_data,

            file_name=download_name,

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            key=(
                "download_"
                +
                str(
                    abs(
                        hash(
                            (
                                filename,
                                sheet_name,
                                invoice_header,
                                header_row,
                                invoice_column,
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
# UPLOAD
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
# APPLICATION
# ============================================================

if uploaded_files:

    st.success(
        f"✅ {len(uploaded_files)} "
        f"Excel file(s) uploaded"
    )

    # --------------------------------------------------------
    # Store uploaded files in memory
    # --------------------------------------------------------

    file_data = tuple(

        (
            file.name,
            file.getvalue()
        )

        for file in uploaded_files
    )

    # --------------------------------------------------------
    # Build index
    # --------------------------------------------------------

    with st.spinner(
        "📖 Reading Excel files..."
    ):

        indexed_files = build_search_index(
            file_data
        )

    # --------------------------------------------------------
    # Errors
    # --------------------------------------------------------

    for file_result in indexed_files:

        if file_result["error"]:

            st.error(
                f"❌ "
                f"{file_result['filename']}: "
                f"{file_result['error']}"
            )

    # --------------------------------------------------------
    # Table count
    # --------------------------------------------------------

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
    # SEARCH AREA
    # ========================================================

    st.markdown(
        "## 🔎 Search Bill / Invoice Number"
    )

    search_value = st.text_input(

        "Bill / Invoice Number",

        placeholder=(
            "Example: 74, 162, "
            "CRSCL/19-20/1086, "
            "SMS/0504/2021-22"
        ),

        label_visibility="collapsed"
    )

    search_button = st.button(

        "🔎 SEARCH",

        type="primary",

        use_container_width=True
    )

    # ========================================================
    # SEARCH
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
                    "Search is performed only in "
                    "detected Bill / Invoice columns."
                )

else:

    st.info(
        "👆 Upload one or more Excel files to begin."
    )
Now your GST/050/18-19 result should show all of this
SNO | BILL NUMBER | SHOP ID | SHOP NAME
    |             |         |
    |             |         |
CAMERA(10) | BNC 20 | ADAPTER10 | DVR 5 | HARDDISK 5

with the actual values:

1 | GST/050/18-19 | 17143 | CHIKKATHOGUR MAIN ROAD BANGALORE
  |               |       |                              | 3970 | 68 | 294 | 3300 | 32

2 | GST/050/18-19 | 17142 | RS NAIDU LAYOUT MYSORE
  |               |       |                      | 3970 | 68 | 294 | 3300 | 32

and so on.

The important difference is that the code now looks through the complete rows belonging to that table to determine the final used column, rather than assuming the table ends at the last non-empty cell in the header row. This preserves all columns through HARDDISK 5.

import streamlit as st
import io
import re
import html
import unicodedata

from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit.components.v1 as components
from openpyxl import load_workbook, Workbook

============================================================
PAGE CONFIG
============================================================

st.set_page_config(
page_title="Bill / Invoice Search",
page_icon="🔎",
layout="wide"
)

============================================================
TITLE
============================================================

st.title("🔎 Bill / Invoice Search")

st.caption(
"Search Bill / Invoice numbers across all uploaded Excel files."
)

============================================================
BASIC FUNCTIONS
============================================================

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

if is_blank(value):
    return ""

# Excel integer
if isinstance(value, int):
    return str(value)

# Excel float
if isinstance(value, float):

    if value.is_integer():
        return str(int(value))

    return str(value).strip()

# Everything else
return str(value).strip()
============================================================
NORMALIZE HEADER
============================================================

def normalize_header(value):

if is_blank(value):
    return ""

text = unicodedata.normalize(
    "NFKC",
    str(value)
)

text = text.strip().lower()

# Remove common separators
text = re.sub(
    r"[_\-/\\.:]+",
    " ",
    text
)

# Remove punctuation
text = re.sub(
    r"[^\w\s]",
    " ",
    text
)

# Multiple spaces
text = re.sub(
    r"\s+",
    " ",
    text
)

return text.strip()
============================================================
BILL / INVOICE HEADER DETECTION


Supports:


BILL
BILL NO
BILL NUMBER
BIL NO
INVOICE
INVOICE NO
INVOICE NUMBER
INOVICE NO <-- typo
INOVICE NUMBER <-- typo
INV
INV NO
INV NUMBER
============================================================

def is_bill_invoice_header(value):

header = normalize_header(value)

if not header:
    return False

# Exact supported names
exact_names = {

    # BILL
    "bill",
    "bill no",
    "bill number",
    "bill num",

    "bil",
    "bil no",
    "bil number",
    "bil num",

    # INVOICE
    "invoice",
    "invoice no",
    "invoice number",
    "invoice num",

    # Common Excel typo
    "inovice",
    "inovice no",
    "inovice number",
    "inovice num",

    # INV
    "inv",
    "inv no",
    "inv number",
    "inv num",

    # Reference formats
    "bill ref",
    "bill ref no",
    "bill ref number",

    "invoice ref",
    "invoice ref no",
    "invoice ref number",

    "inovice ref",
    "inovice ref no",
    "inovice ref number"
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
    "inovice",
    "inv"
}

number_words = {
    "no",
    "number",
    "num"
}

# Example:
#
# BILL NO
# INVOICE NO
# INOVICE NO
# INV NUMBER
#
if (
    words.intersection(bill_words)
    and
    words.intersection(number_words)
):
    return True

return False
============================================================
FIND ALL BILL / INVOICE HEADERS
============================================================

def find_bill_invoice_headers(rows):

headers = []

for row_index, row in enumerate(rows):

    for column_index, value in enumerate(row):

        if is_bill_invoice_header(value):

            headers.append({

                "row": row_index,

                "column": column_index,

                "value": value
            })

return headers
============================================================
DETERMINE TABLE BOUNDARIES


This handles tables like:


S NO | ITEM | PURCHASE DATE | INVOICE NO | SHOP ID ...


and:


S NO | ITEM | PURCHASE DATE | INVOICE | SHOP ID ...


============================================================

def find_table_boundaries(
rows,
header_row,
invoice_column
):

if header_row >= len(rows):

    return (
        invoice_column,
        invoice_column + 1
    )

header = list(
    rows[header_row]
)

column_count = len(header)

# --------------------------------------------------------
# Find first meaningful header to LEFT
# --------------------------------------------------------

start = invoice_column

left = invoice_column - 1

while left >= 0:

    if is_blank(header[left]):
        break

    start = left

    left -= 1

# --------------------------------------------------------
# Find first meaningful header to RIGHT
# --------------------------------------------------------

end = invoice_column

right = invoice_column + 1

while right < column_count:

    if is_blank(header[right]):
        break

    end = right

    right += 1

# --------------------------------------------------------
# Return Python slice
# --------------------------------------------------------

return (
    start,
    end + 1
)
============================================================
FIND END OF THIS TABLE


A table ends when:


1. Another header row for the SAME invoice column appears
2. A new table with a different invoice column/header appears
after sufficient separation


We do NOT stop simply because there is one blank row.
============================================================

def find_table_end(
rows,
header_row,
invoice_column,
all_headers
):

total_rows = len(rows)

possible_next_headers = []

for header in all_headers:

    if header["row"] <= header_row:
        continue

    possible_next_headers.append(
        header
    )

# --------------------------------------------------------
# Look for the next invoice/bill header.
#
# This is the safest boundary because your workbook has
# multiple tables vertically.
# --------------------------------------------------------

for header in sorted(
    possible_next_headers,
    key=lambda x: x["row"]
):

    # If another invoice/bill header occurs later,
    # it represents another table.
    #
    # This is especially important for sheets such as:
    #
    # rows 54-80  -> SMS/0504/2021-22
    # rows 84-89  -> SMS/0976/2021-22
    #
    if header["row"] > header_row:

        return header["row"]

return total_rows
============================================================
BUILD ONE TABLE
============================================================

def build_table(
rows,
filename,
sheet_name,
header_info,
all_headers
):

header_row = header_info[
    "row"
]

invoice_column = header_info[
    "column"
]

invoice_header = header_info[
    "value"
]

# --------------------------------------------------------
# Horizontal range
# --------------------------------------------------------

start_column, end_column = find_table_boundaries(
    rows,
    header_row,
    invoice_column
)

# --------------------------------------------------------
# Vertical range
# --------------------------------------------------------

end_row = find_table_end(
    rows,
    header_row,
    invoice_column,
    all_headers
)

# --------------------------------------------------------
# Get EXACT Excel header names
# --------------------------------------------------------

source_header_row = list(
    rows[header_row]
)

headers = source_header_row[
    start_column:end_column
]

expected_width = (
    end_column -
    start_column
)

while len(headers) < expected_width:

    headers.append(
        None
    )

# --------------------------------------------------------
# Invoice column inside this table
# --------------------------------------------------------

relative_invoice_column = (
    invoice_column -
    start_column
)

# --------------------------------------------------------
# SEARCH INDEX
# --------------------------------------------------------

search_index = defaultdict(list)

# --------------------------------------------------------
# READ ALL ROWS BELONGING TO THIS TABLE
# --------------------------------------------------------

for row_number in range(
    header_row + 1,
    end_row
):

    source_row = list(
        rows[row_number]
    )

    # Make sure row has enough columns
    if len(source_row) < end_column:

        source_row.extend(
            [None] *
            (
                end_column -
                len(source_row)
            )
        )

    # ----------------------------------------------------
    # IMPORTANT:
    #
    # ONLY THIS TABLE'S COLUMNS
    # ----------------------------------------------------

    table_row = source_row[
        start_column:end_column
    ]

    # ----------------------------------------------------
    # Safety
    # ----------------------------------------------------

    if (
        relative_invoice_column < 0
        or
        relative_invoice_column >= len(table_row)
    ):
        continue

    # ----------------------------------------------------
    # Invoice value
    # ----------------------------------------------------

    invoice_value = table_row[
        relative_invoice_column
    ]

    normalized = normalize_value(
        invoice_value
    )

    if normalized == "":
        continue

    # ----------------------------------------------------
    # Store complete table row
    # ----------------------------------------------------

    search_index[
        normalized
    ].append({

        "excel_row":
            row_number + 1,

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
        invoice_header,

    "header_row":
        header_row + 1,

    "invoice_column":
        invoice_column + 1,

    "start_column":
        start_column + 1,

    "end_column":
        end_column,

    "headers":
        headers,

    "search_index":
        dict(search_index)
}
============================================================
READ ONE EXCEL FILE
============================================================

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
    # Every sheet
    # ----------------------------------------------------

    for worksheet in workbook.worksheets:

        rows = []

        # ------------------------------------------------
        # Read complete sheet
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
        # Detect every Bill / Invoice header
        # ------------------------------------------------

        headers = find_bill_invoice_headers(
            rows
        )

        if not headers:
            continue

        # ------------------------------------------------
        # Every header = separate table
        # ------------------------------------------------

        for header_info in headers:

            table = build_table(
                rows,
                filename,
                worksheet.title,
                header_info,
                headers
            )

            if table:

                result[
                    "tables"
                ].append(
                    table
                )

    workbook.close()

except Exception as error:

    result[
        "error"
    ] = str(error)

return result
============================================================
BUILD SEARCH INDEX
============================================================

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

results.sort(
    key=lambda x:
        x["filename"].lower()
)

return results
============================================================
SEARCH
============================================================

def search_excel(
indexed_files,
search_value
):

search_value = normalize_value(
    search_value
)

results = []

seen = set()

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

            unique_key = (

                table["filename"],

                table["sheet"],

                table["header_row"],

                table["invoice_column"],

                match["excel_row"]
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

                "invoice_column":
                    table["invoice_column"],

                "start_column":
                    table["start_column"],

                "end_column":
                    table["end_column"],

                "excel_row":
                    match["excel_row"],

                "headers":
                    table["headers"],

                "row":
                    match["row"]
            })

return results
============================================================
SAFE HTML
============================================================

def safe_html(value):

if is_blank(value):
    return ""

return html.escape(
    str(value)
)
============================================================
DISPLAY TABLE


We intentionally DO NOT use st.dataframe()
because Excel files can contain:


- duplicate headers
- blank headers


which can cause PyArrow errors.
============================================================

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
# Preserve blank Excel headers.
#
# DO NOT rename them to:
#
# Column 1
# Column 2
#
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

<style> html, body { margin: 0; padding: 0; background: transparent; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; } .wrapper { width: 100%; max-height: 650px; overflow-x: auto; overflow-y: auto; border: 1px solid #30333d; border-radius: 8px; } table { border-collapse: collapse; width: max-content; min-width: 100%; table-layout: auto; font-size: 14px; } th { position: sticky; top: 0; z-index: 50; background: #1f2129; color: #e5e7eb; border: 1px solid #3a3d47; padding: 11px 14px; text-align: left; white-space: nowrap; font-weight: 600; min-width: 65px; } td { background: #0e1117; color: #f5f5f5; border: 1px solid #30333d; padding: 10px 14px; text-align: left; white-space: nowrap; min-width: 65px; } tr:hover td { background: #181b23; } ::-webkit-scrollbar { width: 12px; height: 12px; } ::-webkit-scrollbar-track { background: #16181f; } ::-webkit-scrollbar-thumb { background: #777; border-radius: 8px; } ::-webkit-scrollbar-thumb:hover { background: #999; } </style>

</head>

<body>

<div class="wrapper">

<table>

<thead>

<tr> """

# --------------------------------------------------------
# HEADERS
# --------------------------------------------------------

for header in headers:

    table_html += (
        "<th>"
        + safe_html(header)
        + "</th>"
    )

table_html += """

</tr>

</thead>

<tbody> """

# --------------------------------------------------------
# ROWS
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

    table_html += "<tr>"

    for value in row[
        :maximum_columns
    ]:

        table_html += (
            "<td>"
            + safe_html(value)
            + "</td>"
        )

    table_html += "</tr>"

table_html += """

</tbody>

</table>

</div>

</body>

</html> """

calculated_height = min(
    height,
    max(
        220,
        100 + len(rows) * 43
    )
)

components.html(
    table_html,
    height=calculated_height,
    scrolling=False
)
============================================================
DETECTED HEADER TABLE
============================================================

def display_detected_headers(
indexed_files
):

headers = [

    "File",

    "Sheet",

    "Bill / Invoice Header",

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

            table["invoice_column"],

            table["start_column"],

            table["end_column"]
        ])

if rows:

    display_excel_table(
        headers,
        rows,
        height=450
    )
============================================================
DOWNLOAD EXCEL
============================================================

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
# Header row
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
        row[:maximum_columns],
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
# Auto width
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
============================================================
DOWNLOAD FILE NAME
============================================================

def make_download_filename(
filename,
search_value
):

base = re.sub(
    r"\.(xlsx|xlsm)$",
    "",
    filename,
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
============================================================
DISPLAY SEARCH RESULTS
============================================================

def display_results(
results,
search_value
):

# --------------------------------------------------------
# Group matches by ACTUAL TABLE
# --------------------------------------------------------

grouped = defaultdict(list)

for result in results:

    key = (

        result["filename"],

        result["sheet"],

        result["header"],

        result["header_row"],

        result["invoice_column"],

        result["start_column"],

        result["end_column"]
    )

    grouped[
        key
    ].append(
        result
    )

# --------------------------------------------------------
# Display each table
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
        invoice_header,
        header_row,
        invoice_column,
        start_column,
        end_column
    ) = key

    st.markdown(
        "---"
    )

    st.markdown(
        f"### 📄 {filename}"
    )

    st.write(
        f"**Sheet:** {sheet_name}"
    )

    st.write(
        f"**Bill / Invoice Header:** "
        f"`{invoice_header}`"
    )

    st.write(
        f"**Excel Header Row:** "
        f"{header_row}"
        f"  |  "
        f"**Bill / Invoice Column:** "
        f"{invoice_column}"
    )

    st.caption(
        f"Table columns: "
        f"{start_column} → {end_column}"
    )

    st.caption(
        f"Showing {len(group)} "
        f"complete matching Excel row(s)."
    )

    # ----------------------------------------------------
    # EXACT HEADERS
    # ----------------------------------------------------

    headers = list(
        group[0]["headers"]
    )

    # ----------------------------------------------------
    # EXACT MATCHING ROWS
    # ----------------------------------------------------

    rows = [

        result["row"]

        for result in group
    ]

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

    download_name = make_download_filename(
        filename,
        search_value
    )

    st.download_button(

        "⬇️ Download Table",

        data=download_data,

        file_name=download_name,

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
                            invoice_header,
                            header_row,
                            invoice_column,
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
============================================================
UPLOAD
============================================================

uploaded_files = st.file_uploader(

"📂 Upload Excel files",

type=[
    "xlsx",
    "xlsm"
],

accept_multiple_files=True

)

============================================================
APPLICATION
============================================================

if uploaded_files:

st.success(
    f"✅ {len(uploaded_files)} "
    f"Excel file(s) uploaded"
)

# --------------------------------------------------------
# Store uploaded files in memory
# --------------------------------------------------------

file_data = tuple(

    (
        file.name,
        file.getvalue()
    )

    for file in uploaded_files
)

# --------------------------------------------------------
# Build index
# --------------------------------------------------------

with st.spinner(
    "📖 Reading Excel files..."
):

    indexed_files = build_search_index(
        file_data
    )

# --------------------------------------------------------
# Errors
# --------------------------------------------------------

for file_result in indexed_files:

    if file_result["error"]:

        st.error(
            f"❌ "
            f"{file_result['filename']}: "
            f"{file_result['error']}"
        )

# --------------------------------------------------------
# Table count
# --------------------------------------------------------

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
# SEARCH AREA
# ========================================================

st.markdown(
    "## 🔎 Search Bill / Invoice Number"
)

search_value = st.text_input(

    "Bill / Invoice Number",

    placeholder=(
        "Example: 74, 162, "
        "CRSCL/19-20/1086, "
        "SMS/0504/2021-22"
    ),

    label_visibility="collapsed"
)

search_button = st.button(

    "🔎 SEARCH",

    type="primary",

    use_container_width=True
)

# ========================================================
# SEARCH
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
                "Search is performed only in "
                "detected Bill / Invoice columns."
            )

else:

st.info(
    "👆 Upload one or more Excel files to begin."
)
Close
