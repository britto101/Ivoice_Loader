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
# PAGE TITLE
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

        # COMMON TYPO
        "inovice",
        "inovice no",
        "inovice number",
        "inovice num",

        # INV
        "inv",
        "inv no",
        "inv number",
        "inv num",

        # REFERENCE
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
# FIND TABLE WIDTH
#
# IMPORTANT:
#
# Do NOT stop simply because another table has a header.
# We determine the horizontal area of THIS table.
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

    total_columns = len(header)

    # --------------------------------------------------------
    # Start from invoice column and move LEFT.
    #
    # We allow blank headers if the surrounding table contains
    # actual data in that column.
    # --------------------------------------------------------

    start = invoice_column

    # Look up to a few rows around the table.
    # This helps with blank Excel headers.
    sample_end = min(
        len(rows),
        header_row + 8
    )

    for col in range(
        invoice_column - 1,
        -1,
        -1
    ):

        header_value = (
            header[col]
            if col < len(header)
            else None
        )

        has_data = any(
            not is_blank(rows[r][col])
            if col < len(rows[r])
            else False
            for r in range(
                header_row + 1,
                sample_end
            )
        )

        if (
            not is_blank(header_value)
            or has_data
        ):

            start = col

        else:

            # Blank separator column
            break

    # --------------------------------------------------------
    # Move RIGHT.
    # --------------------------------------------------------

    end = invoice_column

    for col in range(
        invoice_column + 1,
        total_columns
    ):

        header_value = (
            header[col]
            if col < len(header)
            else None
        )

        has_data = any(
            not is_blank(rows[r][col])
            if col < len(rows[r])
            else False
            for r in range(
                header_row + 1,
                sample_end
            )
        )

        if (
            not is_blank(header_value)
            or has_data
        ):

            end = col

        else:

            # Blank separator column
            break

    return (
        start,
        end + 1
    )


# ============================================================
# IMPORTANT TABLE OVERLAP CHECK
#
# This is the main fix for your screenshot.
#
# A header from a table on the RIGHT must NOT stop a table
# on the LEFT.
# ============================================================

def horizontal_overlap(
    table_start,
    table_end,
    header_column
):

    return (
        table_start
        <=
        header_column
        <
        table_end
    )


# ============================================================
# FIND TABLE END
#
# ONLY a later header that belongs horizontally inside the
# CURRENT table can terminate the current table.
#
# This prevents:
#
# LEFT TABLE       RIGHT TABLE
# A:H              J:...
#
# from interfering with each other.
# ============================================================

def find_table_end(
    rows,
    header_row,
    table_start,
    table_end,
    invoice_column,
    all_headers
):

    total_rows = len(rows)

    later_headers = sorted(
        [
            h
            for h in all_headers
            if h["row"] > header_row
        ],
        key=lambda x: x["row"]
    )

    for next_header in later_headers:

        next_row = next_header["row"]
        next_column = next_header["column"]

        # ----------------------------------------------------
        # MAIN FIX:
        #
        # Ignore headers belonging to a different horizontal
        # table.
        # ----------------------------------------------------

        if not horizontal_overlap(
            table_start,
            table_end,
            next_column
        ):
            continue

        # ----------------------------------------------------
        # Make sure it is not simply some text that happens
        # to look like a header.
        # ----------------------------------------------------

        # We have a genuine Bill / Invoice header inside the
        # same horizontal table area.
        return next_row

    # --------------------------------------------------------
    # No next table found.
    # --------------------------------------------------------

    return total_rows


# ============================================================
# BUILD TABLE
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
    # Find horizontal table area
    # --------------------------------------------------------

    (
        start_column,
        end_column
    ) = find_table_boundaries(
        rows,
        header_row,
        invoice_column
    )

    # --------------------------------------------------------
    # Find vertical table end.
    #
    # IMPORTANT:
    # Uses horizontal overlap.
    # --------------------------------------------------------

    end_row = find_table_end(
        rows,
        header_row,
        start_column,
        end_column,
        invoice_column,
        all_headers
    )

    # --------------------------------------------------------
    # Header row
    # --------------------------------------------------------

    source_header = list(
        rows[header_row]
    )

    headers = source_header[
        start_column:end_column
    ]

    expected_width = (
        end_column -
        start_column
    )

    while len(headers) < expected_width:

        headers.append(None)

    # --------------------------------------------------------
    # Invoice column relative to table
    # --------------------------------------------------------

    relative_invoice_column = (
        invoice_column -
        start_column
    )

    if (
        relative_invoice_column < 0
        or
        relative_invoice_column >= expected_width
    ):

        return None

    # --------------------------------------------------------
    # Search index
    # --------------------------------------------------------

    search_index = defaultdict(list)

    # --------------------------------------------------------
    # READ EVERY ROW IN TABLE
    # --------------------------------------------------------

    for row_number in range(
        header_row + 1,
        end_row
    ):

        source_row = list(
            rows[row_number]
        )

        # Make row wide enough
        if len(source_row) < end_column:

            source_row.extend(
                [None] *
                (
                    end_column -
                    len(source_row)
                )
            )

        table_row = source_row[
            start_column:end_column
        ]

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
        # Store COMPLETE row
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
    # Return table
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
        # Every worksheet
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
            # Find all Bill / Invoice headers
            # ------------------------------------------------

            headers = find_bill_invoice_headers(
                rows
            )

            if not headers:
                continue

            # ------------------------------------------------
            # Build every table
            # ------------------------------------------------

            for header_info in headers:

                table = build_table(
                    rows,
                    filename,
                    worksheet.title,
                    header_info,
                    headers
                )

                if table is not None:

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
# CACHE EXCEL INDEX
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
# We DO NOT use st.dataframe().
#
# Reason:
# Excel may contain duplicate / blank headers.
# PyArrow can reject duplicate column names.
#
# HTML table avoids that problem completely.
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

    while len(headers) < maximum_columns:

        headers.append(None)

    # --------------------------------------------------------
    # HTML
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
    # EXACT HEADERS
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

<tbody>
"""

    # --------------------------------------------------------
    # ALL ROWS
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
# DETECTED HEADERS
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
# DOWNLOAD RESULT
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

        headers.append(None)

    # --------------------------------------------------------
    # Header
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

    workbook.save(
        output
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# DOWNLOAD NAME
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
# DISPLAY RESULTS
# ============================================================

def display_results(
    results,
    search_value
):

    # --------------------------------------------------------
    # Group by actual table
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
    # Display every table separately
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
        # EXACT ROWS
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
            ),

            use_container_width=False
        )


# ============================================================
# FILE UPLOAD
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

    st.success(
        f"✅ {len(uploaded_files)} "
        f"Excel file(s) uploaded"
    )

    # --------------------------------------------------------
    # Read uploaded files into memory
    # --------------------------------------------------------

    file_data = tuple(

        (
            file.name,
            file.getvalue()
        )

        for file in uploaded_files
    )

    # --------------------------------------------------------
    # Read / index
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
    # Count tables
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
    # SEARCH
    # ========================================================

    st.markdown(
        "## 🔎 Search Bill / Invoice Number"
    )

    search_value = st.text_input(

        "Bill / Invoice Number",

        placeholder=(
            "Example: 74, 162, "
            "321, CRSCL/19-20/1086, "
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
    # PERFORM SEARCH
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
                    "Only detected Bill / Invoice "
                    "columns were searched. "
                    "Matching is exact."
                )

else:

    st.info(
        "👆 Upload one or more Excel files to begin."
    )
