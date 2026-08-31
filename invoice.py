import streamlit as st
import io
import re
import unicodedata
import html

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
    "Search exact Bill / Invoice numbers across all uploaded Excel sheets."
)


# ============================================================
# BASIC HELPERS
# ============================================================

def is_blank(value):
    """
    Check whether an Excel value is blank.
    """

    if value is None:
        return True

    try:
        if value != value:
            return True
    except Exception:
        pass

    return str(value).strip() == ""


# ============================================================
# HEADER NORMALIZATION
#
# Used ONLY for detecting Bill / Invoice headers.
#
# Original Excel header is NEVER changed.
# ============================================================

def normalize_header(value):

    if is_blank(value):
        return ""

    text = unicodedata.normalize(
        "NFKC",
        str(value)
    )

    text = text.strip().lower()

    # Convert common separators into spaces
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

    # Multiple spaces -> one
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

    exact_headers = {

        # BILL
        "bill no",
        "bill number",
        "bill num",

        # BIL
        "bil no",
        "bil number",
        "bil num",

        # INVOICE
        "invoice no",
        "invoice number",
        "invoice num",

        # INV
        "inv no",
        "inv number",
        "inv num",

        # REFERENCE
        "bill ref no",
        "bill ref number",
        "bill reference no",
        "bill reference number",

        "invoice ref no",
        "invoice ref number",
        "invoice reference no",
        "invoice reference number",

        "inv ref no",
        "inv ref number"
    }

    if header in exact_headers:
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

    # Example:
    # Invoice     Number
    # Bill Number
    # BILL-NUMBER
    # invoice_number
    if (
        words.intersection(bill_words)
        and words.intersection(number_words)
    ):
        return True

    return False


# ============================================================
# SEARCH VALUE NORMALIZATION
#
# Used ONLY for exact comparison.
#
# Original Excel value is displayed unchanged.
# ============================================================

def normalize_search_value(value):

    if is_blank(value):
        return ""

    # Excel numeric values
    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

    if isinstance(value, int):
        return str(value)

    return str(value).strip()


# ============================================================
# SAFE DISPLAY VALUE
# ============================================================

def display_value(value):

    if value is None:
        return ""

    try:
        if value != value:
            return ""
    except Exception:
        pass

    return str(value)


# ============================================================
# FIND ACTUAL USED WIDTH
#
# IMPORTANT:
#
# We inspect the actual worksheet rows.
#
# We do NOT use the header row alone because many of your
# Excel files have blank header cells while the data exists
# farther to the right.
# ============================================================

def find_used_width(rows):

    last_column = 0

    for row in rows:

        for column_number, value in enumerate(
            row,
            start=1
        ):

            if not is_blank(value):

                last_column = max(
                    last_column,
                    column_number
                )

    return last_column


# ============================================================
# FIND BILL / INVOICE HEADER ROWS
# ============================================================

def find_bill_headers(rows):

    found = []

    for row_index, row in enumerate(
        rows
    ):

        for column_index, value in enumerate(
            row
        ):

            if is_bill_invoice_header(
                value
            ):

                found.append({

                    "row":
                        row_index,

                    "column":
                        column_index,

                    "value":
                        value
                })

    return found


# ============================================================
# BUILD SEARCH INDEX
#
# IMPORTANT CHANGE:
#
# We search the COMPLETE sheet after finding a Bill/Invoice
# header.
#
# We DO NOT stop because of blank rows.
#
# This fixes the problem where 20 matching rows were reduced
# to 13.
# ============================================================

def build_sheet_index(
    sheet_name,
    rows,
    filename
):

    if not rows:
        return []

    # --------------------------------------------------------
    # Find all Bill / Invoice headers
    # --------------------------------------------------------

    detected_headers = find_bill_headers(
        rows
    )

    if not detected_headers:
        return []

    # --------------------------------------------------------
    # Find complete used width of entire sheet
    #
    # This makes sure columns after TOTAL are included.
    # --------------------------------------------------------

    used_width = find_used_width(
        rows
    )

    if used_width == 0:
        return []

    results = []

    # --------------------------------------------------------
    # Group header occurrences by column.
    #
    # If the same Bill/Invoice header appears again lower
    # in the sheet, it is treated as another section.
    # --------------------------------------------------------

    headers_by_column = defaultdict(list)

    for header in detected_headers:

        headers_by_column[
            header["column"]
        ].append(
            header
        )

    # --------------------------------------------------------
    # Process every detected Bill/Invoice column
    # --------------------------------------------------------

    for bill_column, header_list in headers_by_column.items():

        # ----------------------------------------------------
        # If there are repeated header rows in the same column,
        # each section starts at its own header row.
        # ----------------------------------------------------

        for header_position, header_info in enumerate(
            header_list
        ):

            header_row = header_info[
                "row"
            ]

            bill_header = header_info[
                "value"
            ]

            # ------------------------------------------------
            # IMPORTANT:
            #
            # We no longer stop at blank rows.
            #
            # Search continues until:
            #
            # 1. Another Bill/Invoice header in the same
            #    column is reached
            #
            # OR
            #
            # 2. End of worksheet.
            #
            # ------------------------------------------------

            end_row = len(
                rows
            )

            if (
                header_position + 1
                < len(header_list)
            ):

                end_row = header_list[
                    header_position + 1
                ]["row"]

            # ------------------------------------------------
            # Actual Excel header row ONLY
            #
            # DO NOT search upward.
            #
            # This fixes the problem where values such as
            # "9", "DIGISOL MODEM", etc. were being used
            # as column names.
            # ------------------------------------------------

            original_headers = list(
                rows[
                    header_row
                ][:used_width]
            )

            # Make sure header list has complete width
            if len(original_headers) < used_width:

                original_headers.extend(
                    [None] *
                    (
                        used_width -
                        len(original_headers)
                    )
                )

            # ------------------------------------------------
            # Search map
            # ------------------------------------------------

            search_map = defaultdict(list)

            # ------------------------------------------------
            # Search every data row
            #
            # NO blank-row termination.
            # ------------------------------------------------

            for row_number in range(
                header_row + 1,
                end_row
            ):

                row = list(
                    rows[
                        row_number
                    ]
                )

                # ------------------------------------------------
                # Make complete row
                # ------------------------------------------------

                if len(row) < used_width:

                    row.extend(
                        [None] *
                        (
                            used_width -
                            len(row)
                        )
                    )

                # Keep only actual used width
                row = row[
                    :used_width
                ]

                # ------------------------------------------------
                # Bill / Invoice value
                # ------------------------------------------------

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

                # ------------------------------------------------
                # Store row number + complete row
                # ------------------------------------------------

                search_map[
                    search_value
                ].append({

                    "excel_row":
                        row_number + 1,

                    "row":
                        row
                })

            # ------------------------------------------------
            # Store table
            # ------------------------------------------------

            results.append({

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

                "headers":
                    original_headers,

                "search_map":
                    dict(search_map),

                "used_width":
                    used_width
            })

    return results


# ============================================================
# READ ONE EXCEL FILE
#
# openpyxl read_only=True is used for speed and lower memory.
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
        # Load workbook
        # ----------------------------------------------------

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

            # ------------------------------------------------
            # Read complete worksheet
            #
            # No pandas.
            # No header guessing.
            # ------------------------------------------------

            rows = []

            for row in worksheet.iter_rows(
                values_only=True
            ):

                rows.append(
                    list(row)
                )

            if not rows:
                continue

            # ------------------------------------------------
            # Build index
            # ------------------------------------------------

            tables = build_sheet_index(
                worksheet.title,
                rows,
                filename
            )

            result[
                "tables"
            ].extend(
                tables
            )

        workbook.close()

    except Exception as e:

        result["error"] = str(
            e
        )

    return result


# ============================================================
# BUILD COMPLETE SEARCH INDEX
# ============================================================

@st.cache_data(
    show_spinner=False,
    max_entries=5
)
def build_search_index(
    file_data
):

    results = []

    # --------------------------------------------------------
    # Multiple uploaded files can be processed in parallel.
    # --------------------------------------------------------

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
    # Keep file order
    # --------------------------------------------------------

    results.sort(
        key=lambda item:
            item["filename"].lower()
    )

    return results


# ============================================================
# SEARCH ALL INDEXED BILL / INVOICE COLUMNS
# ============================================================

def search_excel(
    indexed_files,
    search_value
):

    search_value = normalize_search_value(
        search_value
    )

    results = []

    # Used to avoid duplicate results
    # if Excel has repeated header detection.
    seen = set()

    for file_data in indexed_files:

        for table in file_data[
            "tables"
        ]:

            matches = table[
                "search_map"
            ].get(
                search_value,
                []
            )

            for match in matches:

                excel_row = match[
                    "excel_row"
                ]

                row = match[
                    "row"
                ]

                unique_key = (
                    table["filename"],
                    table["sheet"],
                    excel_row,
                    table["bill_column"]
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

                    "excel_row":
                        excel_row,

                    "headers":
                        table["headers"],

                    "row":
                        row
                })

    return results


# ============================================================
# DETECTED HEADER DISPLAY
# ============================================================

def show_detected_headers(
    indexed_files
):

    detected = []

    for file_data in indexed_files:

        for table in file_data[
            "tables"
        ]:

            detected.append({

                "File":
                    table["filename"],

                "Sheet":
                    table["sheet"],

                "Header":
                    table["header"],

                "Excel Header Row":
                    table["header_row"],

                "Bill / Invoice Column":
                    table["bill_column"],

                "Total Columns":
                    table["used_width"]
            })

    if detected:

        # Use HTML table instead of dataframe so even
        # unusual values cannot cause duplicate-column errors.

        headers = [
            "File",
            "Sheet",
            "Header",
            "Excel Header Row",
            "Bill / Invoice Column",
            "Total Columns"
        ]

        rows = []

        for item in detected:

            rows.append([
                item["File"],
                item["Sheet"],
                item["Header"],
                item["Excel Header Row"],
                item["Bill / Invoice Column"],
                item["Total Columns"]
            ])

        display_excel_table(
            headers,
            rows,
            max_height=450
        )


# ============================================================
# HTML ESCAPE
# ============================================================

def html_value(value):

    return html.escape(
        display_value(value)
    )


# ============================================================
# DISPLAY TABLE
#
# IMPORTANT:
#
# We do NOT use st.dataframe().
#
# This supports:
#
# - duplicate headers
# - blank headers
# - wide Excel sheets
# - all columns
# ============================================================

def display_excel_table(
    headers,
    rows,
    max_height=650
):

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
    # Ensure header count matches data width
    #
    # Blank headers remain blank.
    # --------------------------------------------------------

    while len(headers) < maximum_columns:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # Build HTML
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

.table-wrapper {

    width: 100%;

    max-height: 650px;

    overflow-x: auto;

    overflow-y: auto;

    border:
        1px solid #30333d;

    border-radius:
        8px;

    scrollbar-width: auto;
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

thead th {

    position:
        sticky;

    top:
        0;

    z-index:
        20;

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

    min-width:
        60px;
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

    min-width:
        60px;
}

tbody tr:hover td {

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

<div class="table-wrapper">

<table>

<thead>

<tr>
"""

    # ========================================================
    # HEADER
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
    # DATA
    # ========================================================

    for row in rows:

        row = list(
            row
        )

        # Complete row
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
    # Height
    # --------------------------------------------------------

    calculated_height = min(
        max_height,
        100 + (
            len(rows) * 45
        )
    )

    calculated_height = max(
        180,
        calculated_height
    )

    # --------------------------------------------------------
    # Render
    # --------------------------------------------------------

    components.html(
        table_html,
        height=calculated_height,
        scrolling=False
    )


# ============================================================
# CREATE DOWNLOAD EXCEL
#
# Downloads ONLY the matching rows.
#
# It includes ALL columns displayed in the result.
#
# Duplicate / blank headers are preserved as far as Excel
# itself allows.
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
    # Determine complete width
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
    # Complete header width
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

        if is_blank(header):

            cell_value = ""

        else:

            cell_value = header

        worksheet.cell(
            row=1,
            column=column_number,
            value=cell_value
        )

    # ========================================================
    # WRITE ROWS
    # ========================================================

    for excel_row_number, row in enumerate(
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

            if is_blank(value):

                cell_value = None

            else:

                cell_value = value

            worksheet.cell(
                row=excel_row_number,
                column=column_number,
                value=cell_value
            )

    # ========================================================
    # FREEZE HEADER
    # ========================================================

    worksheet.freeze_panes = "A2"

    # ========================================================
    # COLUMN WIDTH
    # ========================================================

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
# DOWNLOAD FILE NAME
# ============================================================

def make_download_filename(
    filename,
    search_value
):

    base_name = re.sub(
        r"\.(xlsx|xlsm)$",
        "",
        filename,
        flags=re.IGNORECASE
    )

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
    # Group by file / sheet / header
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

        grouped[
            key
        ].append(
            result
        )

    # --------------------------------------------------------
    # Display groups
    # --------------------------------------------------------

    for group_index, (
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
            bill_column
        ) = key

        st.markdown(
            "---"
        )

        # ----------------------------------------------------
        # FILE
        # ----------------------------------------------------

        st.markdown(
            f"### 📄 {filename}"
        )

        # ----------------------------------------------------
        # SHEET
        # ----------------------------------------------------

        st.write(
            f"**Sheet:** {sheet_name}"
        )

        # ----------------------------------------------------
        # BILL / INVOICE HEADER
        # ----------------------------------------------------

        st.write(
            f"**Bill / Invoice Header:** "
            f"`{bill_header}`"
        )

        # ----------------------------------------------------
        # HEADER INFORMATION
        # ----------------------------------------------------

        st.write(
            f"**Excel Header Row:** "
            f"{header_row}"
            f"  |  "
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
        # Complete matching rows
        # ----------------------------------------------------

        rows = [
            result["row"]
            for result in group
        ]

        # ----------------------------------------------------
        # Display Excel row information
        # ----------------------------------------------------

        st.caption(
            f"Showing {len(rows)} complete matching "
            f"Excel row(s)."
        )

        # ----------------------------------------------------
        # DISPLAY COMPLETE TABLE
        # ----------------------------------------------------

        display_excel_table(
            headers,
            rows,
            max_height=650
        )

        # ====================================================
        # DOWNLOAD
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
        # Unique key
        # ----------------------------------------------------

        download_key = (
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
                            len(rows),
                            group_index
                        )
                    )
                )
            )
        )

        st.download_button(

            label="⬇️ Download Table",

            data=download_data,

            file_name=download_filename,

            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            key=download_key
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

    # ========================================================
    # UPLOAD STATUS
    # ========================================================

    st.success(
        f"✅ {len(uploaded_files)} "
        f"Excel file(s) uploaded"
    )

    # ========================================================
    # STORE FILES IN MEMORY
    # ========================================================

    file_data = tuple(

        (
            file.name,
            file.getvalue()
        )

        for file in uploaded_files
    )

    # ========================================================
    # BUILD INDEX
    # ========================================================

    with st.spinner(
        "📖 Reading Excel files and building search index..."
    ):

        indexed_files = build_search_index(
            file_data
        )

    # ========================================================
    # SHOW ERRORS
    # ========================================================

    for file_result in indexed_files:

        if file_result["error"]:

            st.error(
                f"❌ {file_result['filename']}: "
                f"{file_result['error']}"
            )

    # ========================================================
    # COUNT TABLES
    # ========================================================

    total_tables = sum(

        len(
            file_result["tables"]
        )

        for file_result in indexed_files
    )

    # ========================================================
    # DETECTED HEADER STATUS
    # ========================================================

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
    # SEARCH SECTION
    # ========================================================

    st.markdown(
        "## 🔎 Search Bill / Invoice Number"
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
        # No headers
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
