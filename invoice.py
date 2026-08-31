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
    "Search exact Bill / Invoice numbers across uploaded Excel files."
)


# ============================================================
# SETTINGS
# ============================================================

# Maximum number of rows allowed between a header row and
# another possible header row.
HEADER_LOOKBACK = 10

# Maximum number of completely blank rows used to identify
# the end of a table section.
MAX_EMPTY_ROWS = 2


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


def normalize_search_value(value):

    if is_blank(value):
        return ""

    if isinstance(value, float):

        if value.is_integer():
            return str(int(value))

    if isinstance(value, int):
        return str(value)

    return str(value).strip()


def display_value(value):

    if is_blank(value):
        return ""

    return str(value)


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

    # Convert separators to spaces
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

    exact_headers = {

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

    if (
        words.intersection(bill_words)
        and words.intersection(number_words)
    ):
        return True

    return False


# ============================================================
# FIND BILL / INVOICE HEADERS
# ============================================================

def find_bill_headers(rows):

    found = []

    for row_index, row in enumerate(rows):

        for column_index, value in enumerate(row):

            if is_bill_invoice_header(value):

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
# FIND TABLE COLUMN RANGE
#
# THIS IS THE IMPORTANT FIX.
#
# Your worksheet contains multiple tables SIDE-BY-SIDE.
#
# Example:
#
# LEFT TABLE:
# S NO | ITEM | INVOICE NO | SHOP ID | ...
#
# RIGHT TABLE:
# S NO | ITEM | INVOICE NO | SHOP ID | ...
#
# We must return ONLY the table containing the searched
# Bill / Invoice header.
#
# We therefore determine the table's horizontal boundaries
# from the header row.
# ============================================================

def find_table_column_range(
    rows,
    header_row,
    bill_column
):

    if header_row >= len(rows):
        return bill_column, bill_column + 1

    row = rows[
        header_row
    ]

    total_columns = len(row)

    # --------------------------------------------------------
    # Find contiguous header block around Bill/Invoice column
    # --------------------------------------------------------

    start_column = bill_column
    end_column = bill_column

    # --------------------------------------------------------
    # Search LEFT
    #
    # Continue while header cells exist.
    # --------------------------------------------------------

    column = bill_column - 1

    while column >= 0:

        value = row[column]

        if is_blank(value):
            break

        start_column = column

        column -= 1

    # --------------------------------------------------------
    # Search RIGHT
    # --------------------------------------------------------

    column = bill_column + 1

    while column < total_columns:

        value = row[column]

        if is_blank(value):
            break

        end_column = column

        column += 1

    # --------------------------------------------------------
    # If only Bill/Invoice itself was detected, try to find
    # surrounding columns using nearby rows.
    #
    # This helps with some irregular Excel sheets.
    # --------------------------------------------------------

    if start_column == bill_column and end_column == bill_column:

        # Look a few rows above for a possible header layout
        for previous_row in range(
            max(
                0,
                header_row - HEADER_LOOKBACK
            ),
            header_row
        ):

            candidate = rows[
                previous_row
            ]

            if bill_column >= len(candidate):
                continue

            # Only use it if the Bill/Invoice position is
            # also occupied or nearby.
            left = bill_column

            while left > 0:

                if is_blank(
                    candidate[left - 1]
                ):
                    break

                left -= 1

            right = bill_column

            while right + 1 < len(candidate):

                if is_blank(
                    candidate[right + 1]
                ):
                    break

                right += 1

            if (
                left != bill_column
                or right != bill_column
            ):

                start_column = left
                end_column = right

                break

    return (
        start_column,
        end_column + 1
    )


# ============================================================
# FIND TABLE END
#
# IMPORTANT:
#
# We search until another Bill/Invoice header appears IN THE
# SAME TABLE COLUMN.
#
# We do not treat every blank row in the worksheet as the end.
# ============================================================

def find_table_end(
    rows,
    header_row,
    bill_column,
    all_headers
):

    total_rows = len(rows)

    end_row = total_rows

    for header in all_headers:

        if header["column"] != bill_column:
            continue

        if header["row"] <= header_row:
            continue

        end_row = header[
            "row"
        ]

        break

    return end_row


# ============================================================
# REMOVE COMPLETELY EMPTY TRAILING COLUMNS
#
# We keep the table's actual header columns.
#
# We do NOT use the entire worksheet.
# ============================================================

def trim_table_columns(
    rows,
    header_row,
    start_column,
    end_column,
    end_row
):

    if end_column <= start_column:
        return start_column, end_column

    # --------------------------------------------------------
    # Header itself defines the initial range.
    # --------------------------------------------------------

    actual_start = start_column
    actual_end = end_column

    # --------------------------------------------------------
    # Do not shrink columns that have actual header names.
    #
    # This preserves the Excel table structure.
    # --------------------------------------------------------

    return actual_start, actual_end


# ============================================================
# BUILD ONE TABLE INDEX
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

    start_column, end_column = find_table_column_range(
        rows,
        header_row,
        bill_column
    )

    # --------------------------------------------------------
    # Find vertical end
    # --------------------------------------------------------

    end_row = find_table_end(
        rows,
        header_row,
        bill_column,
        all_headers
    )

    # --------------------------------------------------------
    # Headers ONLY from this table
    # --------------------------------------------------------

    headers = list(
        rows[
            header_row
        ][
            start_column:end_column
        ]
    )

    # --------------------------------------------------------
    # Ensure header width
    # --------------------------------------------------------

    table_width = (
        end_column -
        start_column
    )

    if len(headers) < table_width:

        headers.extend(
            [None] *
            (
                table_width -
                len(headers)
            )
        )

    # --------------------------------------------------------
    # Search index
    # --------------------------------------------------------

    search_map = defaultdict(list)

    # --------------------------------------------------------
    # Search all rows belonging to this section
    # --------------------------------------------------------

    for row_number in range(
        header_row + 1,
        end_row
    ):

        original_row = list(
            rows[
                row_number
            ]
        )

        # ----------------------------------------------------
        # Ensure enough columns
        # ----------------------------------------------------

        if len(original_row) < end_column:

            original_row.extend(
                [None] *
                (
                    end_column -
                    len(original_row)
                )
            )

        # ----------------------------------------------------
        # ONLY THIS TABLE'S COLUMNS
        # ----------------------------------------------------

        table_row = original_row[
            start_column:end_column
        ]

        # ----------------------------------------------------
        # Bill/Invoice position relative to table
        # ----------------------------------------------------

        relative_bill_column = (
            bill_column -
            start_column
        )

        if (
            relative_bill_column < 0
            or relative_bill_column >= len(table_row)
        ):
            continue

        bill_value = table_row[
            relative_bill_column
        ]

        search_value = normalize_search_value(
            bill_value
        )

        if not search_value:
            continue

        # ----------------------------------------------------
        # Store COMPLETE TABLE ROW
        #
        # NOT the entire worksheet row.
        # ----------------------------------------------------

        search_map[
            search_value
        ].append({

            "excel_row":
                row_number + 1,

            "row":
                table_row
        })

    # --------------------------------------------------------
    # Return table index
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

        "search_map":
            dict(search_map)
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
        # Every worksheet
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

            bill_headers = find_bill_headers(
                rows
            )

            if not bill_headers:
                continue

            # ------------------------------------------------
            # IMPORTANT:
            #
            # Each detected header creates its own table.
            #
            # Tables can exist side-by-side.
            # ------------------------------------------------

            for header_info in bill_headers:

                table = build_table_index(
                    rows,
                    filename,
                    worksheet.title,
                    header_info,
                    bill_headers
                )

                if table:

                    result[
                        "tables"
                    ].append(
                        table
                    )

        workbook.close()

    except Exception as e:

        result["error"] = str(
            e
        )

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

    results.sort(
        key=lambda item:
            item["filename"].lower()
    )

    return results


# ============================================================
# SEARCH
# ============================================================

def search_excel(
    indexed_files,
    search_value
):

    search_value = normalize_search_value(
        search_value
    )

    results = []

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

                    "excel_row":
                        excel_row,

                    "headers":
                        table["headers"],

                    "row":
                        match["row"]
                })

    return results


# ============================================================
# HTML VALUE
# ============================================================

def html_value(value):

    return html.escape(
        display_value(value)
    )


# ============================================================
# DISPLAY HTML TABLE
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
    # Determine width
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

    components.html(
        table_html,
        height=calculated_height,
        scrolling=False
    )


# ============================================================
# DETECTED HEADER DISPLAY
# ============================================================

def show_detected_headers(
    indexed_files
):

    headers = [
        "File",
        "Sheet",
        "Header",
        "Excel Header Row",
        "Bill / Invoice Column",
        "Table Start Column",
        "Table End Column"
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
            max_height=450
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

    # --------------------------------------------------------
    # Determine width
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
    # Complete headers
    # --------------------------------------------------------

    while len(headers) < maximum_columns:

        headers.append(
            None
        )

    # --------------------------------------------------------
    # Write headers
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
    # Write rows
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
    # Column widths
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
# DOWNLOAD FILENAME
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
    # Group results by actual table
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
    # Display each table
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
        # HEADER INFO
        # ----------------------------------------------------

        st.write(
            f"**Excel Header Row:** "
            f"{header_row}"
            f"  |  "
            f"**Bill / Invoice Column:** "
            f"{bill_column}"
        )

        # ----------------------------------------------------
        # MATCH COUNT
        # ----------------------------------------------------

        st.caption(
            f"Showing {len(group)} "
            f"complete matching Excel row(s)."
        )

        # ----------------------------------------------------
        # HEADERS FROM THIS TABLE ONLY
        # ----------------------------------------------------

        headers = list(
            group[0]["headers"]
        )

        # ----------------------------------------------------
        # MATCHING ROWS FROM THIS TABLE ONLY
        # ----------------------------------------------------

        rows = [
            item["row"]
            for item in group
        ]

        # ----------------------------------------------------
        # DISPLAY
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
    # KEEP FILES IN MEMORY
    # ========================================================

    file_data = tuple(

        (
            file.name,
            file.getvalue()
        )

        for file in uploaded_files
    )

    # ========================================================
    # READ FILES
    # ========================================================

    with st.spinner(
        "📖 Reading Excel files and building search index..."
    ):

        indexed_files = build_search_index(
            file_data
        )

    # ========================================================
    # ERRORS
    # ========================================================

    for file_result in indexed_files:

        if file_result["error"]:

            st.error(
                f"❌ {file_result['filename']}: "
                f"{file_result['error']}"
            )

    # ========================================================
    # TOTAL TABLES
    # ========================================================

    total_tables = sum(

        len(
            file_result["tables"]
        )

        for file_result in indexed_files
    )

    # ========================================================
    # DETECTED HEADERS
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
    # SEARCH
    # ========================================================

    st.markdown(
        "## 🔎 Search Bill / Invoice Number"
    )

    search_value = st.text_input(

        "Bill / Invoice Number",

        placeholder="Example: CRSCL/19-20/1086",

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
        # Empty input
        # ----------------------------------------------------

        if not search_value.strip():

            st.warning(
                "⚠️ Please enter a "
                "Bill / Invoice number."
            )

        # ----------------------------------------------------
        # No tables
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
                    "Only detected Bill / Invoice "
                    "columns were searched. "
                    "Matching is exact."
                )

else:

    st.info(
        "👆 Upload one or more Excel files to begin."
    )
