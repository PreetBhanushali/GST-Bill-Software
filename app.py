
from pathlib import Path
from datetime import date, datetime
import calendar
import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ==================================================
# 1. APPLICATION CONFIGURATION
# ==================================================

st.set_page_config(
    page_title="GST Sales Management",
    page_icon="🧾",
    layout="wide",
)

APP_DIR = Path(__file__).resolve().parent

COLUMNS = [
    "Date",
    "Bill No.",
    "Rupees (₹)",
    "GST Credit Party Name",
    "Without GST (₹)",
    "CGST 9% (₹)",
    "SGST 9% (₹)",
    "With Final Total (₹)",
]

AMOUNT_COLUMNS = [
    "Rupees (₹)",
    "Without GST (₹)",
    "CGST 9% (₹)",
    "SGST 9% (₹)",
    "With Final Total (₹)",
]

# Default filename = previous month's sales
today = date.today()
previous_month = today.month - 1 or 12
previous_year = today.year if today.month > 1 else today.year - 1

default_filename = (
    f"GST SALES {calendar.month_abbr[previous_month].upper()} "
    f"{previous_year}"
)

if "active_workbook_name" not in st.session_state:
    st.session_state["active_workbook_name"] = default_filename


# ==================================================
# 2. MONTHLY EXCEL FILE SELECTOR
# ==================================================

st.sidebar.title("Monthly Excel File")

with st.sidebar.form("workbook_selector"):
    requested_name = st.text_input(
        "Excel filename (without .xlsx)",
        value=st.session_state["active_workbook_name"],
        placeholder="GST SALES SEP 2026",
    )

    open_file_clicked = st.form_submit_button(
        "Create / Open File",
        type="primary",
        use_container_width=True,
    )

if open_file_clicked:
    clean_name = re.sub(
        r'[<>:"/\\|?*]',
        "",
        requested_name.strip(),
    ).strip().rstrip(".")

    if not clean_name or clean_name in {".", ".."}:
        st.sidebar.error("Enter a valid filename.")
    else:
        st.session_state["active_workbook_name"] = clean_name

EXCEL_FILE = (
    APP_DIR / f"{st.session_state['active_workbook_name']}.xlsx"
)

st.sidebar.caption(f"Current file: {EXCEL_FILE.name}")


# ==================================================
# 3. EXCEL WORKBOOK FUNCTIONS
# ==================================================

def format_worksheet(ws):
    ws.freeze_panes = "A2"

    widths = [15, 18, 17, 32, 20, 18, 18, 23]

    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    for cell in ws[1]:
        cell.font = Font(
            bold=True,
            color="FFFFFF",
            size=11,
        )
        cell.fill = PatternFill(
            "solid",
            fgColor="17365D",
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    ws.row_dimensions[1].height = 28
    ws.auto_filter.ref = ws.dimensions

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 1).number_format = "dd-mm-yyyy"

        for col in (3, 5, 6, 7, 8):
            ws.cell(row, col).number_format = '#,##0.00'


def ensure_workbook():
    """Create a new workbook only if the selected file is absent."""

    if EXCEL_FILE.exists():
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Bill Records"
    ws.append(COLUMNS)

    format_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def load_records():
    """Load records and preserve actual Excel row numbers."""

    ensure_workbook()

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Bill Records"]

    records = []

    for row_num in range(2, ws.max_row + 1):
        values = [
            ws.cell(row_num, col).value
            for col in range(1, len(COLUMNS) + 1)
        ]

        if all(value is None for value in values):
            continue

        record = dict(zip(COLUMNS, values))
        record["_excel_row"] = row_num
        records.append(record)

    wb.close()
    return records


def save_record(record):
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    ws.append([record.get(column) for column in COLUMNS])

    format_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def update_record(row_num, record):
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    for col_num, column in enumerate(COLUMNS, start=1):
        ws.cell(row_num, col_num).value = record.get(column)

    format_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def delete_record(row_num):
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    ws.delete_rows(row_num, 1)

    format_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def make_record(
    bill_date,
    bill_no,
    rupees,
    party,
    without_gst,
    cgst,
    sgst,
    final_total,
):
    return {
        "Date": datetime.combine(
            bill_date, datetime.min.time()
        ),
        "Bill No.": bill_no.strip(),
        "Rupees (₹)": rupees,
        "GST Credit Party Name": party.strip(),
        "Without GST (₹)": without_gst,
        "CGST 9% (₹)": cgst,
        "SGST 9% (₹)": sgst,
        "With Final Total (₹)": final_total,
    }


def numeric_value(value):
    """Convert an empty input into a blank Excel cell."""

    if value is None:
        return None

    return float(value)


def make_excel_download():
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    format_worksheet(ws)

    output = BytesIO()
    wb.save(output)
    wb.close()

    output.seek(0)
    return output.getvalue()


def format_date(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%d-%m-%Y")
    return str(value or "")


# Create or open the selected monthly workbook
try:
    ensure_workbook()
except PermissionError:
    st.error(
        "The Excel file is open or locked. "
        "Close it in Excel and restart the operation."
    )
    st.stop()


# ==================================================
# 4. HEADER AND SUMMARY
# ==================================================

st.title("🧾 GST Sales Management System")
st.caption(
    "Enter and manage monthly sales bills using Excel."
)

records = load_records()

total_bills = len(records)

total_sales = sum(
    float(record["With Final Total (₹)"] or 0)
    for record in records
)

col1, col2, col3 = st.columns(3)

col1.metric("Total Bills", total_bills)
col2.metric("Total Sales", f"₹ {total_sales:,.2f}")
col3.metric("Active Excel File", EXCEL_FILE.name)

st.divider()


# ==================================================
# 5. APPLICATION TABS
# ==================================================

tab_add, tab_view, tab_edit, tab_export = st.tabs(
    [
        "➕ Add Bill",
        "📋 View & Search",
        "✏️ Edit / Delete",
        "📥 Download Excel",
    ]
)


# ==================================================
# 6. ADD BILL
# ==================================================

with tab_add:
    st.subheader("Enter New Bill")

    with st.form("add_bill_form", clear_on_submit=True):

        left, right = st.columns(2)

        with left:
            bill_date = st.date_input(
                "1. Date",
                value=date.today(),
            )

            bill_no = st.text_input(
                "2. Bill No.",
                placeholder="Enter bill number",
            )

            rupees = st.number_input(
                "3. Rupees (₹)",
                min_value=0.0,
                value=None,
                step=100.0,
                format="%.2f",
                help="Optional numeric amount.",
            )

            party = st.text_input(
                "4. GST Credit Party Name (Optional)",
                placeholder="Enter party name, if applicable",
            )

        with right:
            without_gst = st.number_input(
                "5. Without GST (₹)",
                min_value=0.0,
                value=None,
                step=100.0,
                format="%.2f",
            )

            cgst = st.number_input(
                "6. CGST 9% (₹)",
                min_value=0.0,
                value=None,
                step=10.0,
                format="%.2f",
                help="Enter the CGST amount manually.",
            )

            sgst = st.number_input(
                "7. SGST 9% (₹)",
                min_value=0.0,
                value=None,
                step=10.0,
                format="%.2f",
                help="Enter the SGST amount manually.",
            )

            final_total = st.number_input(
                "8. With Final Total (₹)",
                min_value=0.0,
                value=None,
                step=100.0,
                format="%.2f",
                help="Enter the final total manually.",
            )

        save_clicked = st.form_submit_button(
            "💾 Save Bill",
            type="primary",
            use_container_width=True,
        )

    if save_clicked:
        if not bill_no.strip():
            st.error("Bill No. is required.")
        else:
            existing = load_records()

            duplicate = any(
                str(r["Bill No."] or "").strip().casefold()
                == bill_no.strip().casefold()
                for r in existing
            )

            if duplicate:
                st.error(
                    "This Bill No. already exists in this workbook."
                )
            else:
                record = make_record(
                    bill_date,
                    bill_no,
                    numeric_value(rupees),
                    party,
                    numeric_value(without_gst),
                    numeric_value(cgst),
                    numeric_value(sgst),
                    numeric_value(final_total),
                )

                try:
                    save_record(record)
                    st.success("Bill saved successfully.")
                    st.rerun()
                except PermissionError:
                    st.error(
                        "Close the Excel file and try again."
                    )


# ==================================================
# 7. VIEW AND SEARCH
# ==================================================

with tab_view:
    st.subheader("Saved Bill Records")

    search_term = st.text_input(
        "Search by Bill No. or Party Name",
        placeholder="Type to search...",
        key="record_search",
    )

    records = load_records()

    if search_term.strip():
        term = search_term.strip().casefold()

        records = [
            r for r in records
            if term in str(r["Bill No."] or "").casefold()
            or term in str(
                r["GST Credit Party Name"] or ""
            ).casefold()
        ]

    if records:
        display_df = pd.DataFrame([
            {column: record[column] for column in COLUMNS}
            for record in records
        ])

        display_df["Date"] = display_df["Date"].apply(format_date)

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
        )

        st.caption(f"{len(display_df)} record(s) displayed.")
    else:
        st.info("No matching records found.")


# ==================================================
# 8. EDIT OR DELETE
# ==================================================

with tab_edit:
    st.subheader("Edit / Delete Existing Bill")

    records = load_records()

    if not records:
        st.info("There are no records to edit.")
    else:
        def record_label(record):
            return (
                f"Row {record['_excel_row']} | "
                f"{record['Bill No.']} | "
                f"{record['GST Credit Party Name'] or '(No party name)'}"
            )

        labels = [record_label(r) for r in records]

        selected_label = st.selectbox(
            "Select bill",
            options=labels,
        )

        selected = next(
            r for r in records
            if record_label(r) == selected_label
        )

        row_num = selected["_excel_row"]

        stored_date = selected["Date"]
        if isinstance(stored_date, datetime):
            stored_date = stored_date.date()
        elif not isinstance(stored_date, date):
            stored_date = date.today()

        def stored_amount(column):
            value = selected[column]
            return float(value) if value is not None else None

        with st.form("edit_bill_form"):
            edit_date = st.date_input(
                "Date",
                value=stored_date,
            )

            edit_bill_no = st.text_input(
                "Bill No.",
                value=str(selected["Bill No."] or ""),
            )

            edit_rupees = st.number_input(
                "Rupees (₹)",
                min_value=0.0,
                value=stored_amount("Rupees (₹)"),
                step=100.0,
                format="%.2f",
            )

            edit_party = st.text_input(
                "GST Credit Party Name (Optional)",
                value=str(selected["GST Credit Party Name"] or ""),
            )

            edit_without = st.number_input(
                "Without GST (₹)",
                min_value=0.0,
                value=stored_amount("Without GST (₹)"),
                step=100.0,
                format="%.2f",
            )

            edit_cgst = st.number_input(
                "CGST 9% (₹)",
                min_value=0.0,
                value=stored_amount("CGST 9% (₹)"),
                step=10.0,
                format="%.2f",
            )

            edit_sgst = st.number_input(
                "SGST 9% (₹)",
                min_value=0.0,
                value=stored_amount("SGST 9% (₹)"),
                step=10.0,
                format="%.2f",
            )

            edit_total = st.number_input(
                "With Final Total (₹)",
                min_value=0.0,
                value=stored_amount("With Final Total (₹)"),
                step=100.0,
                format="%.2f",
            )

            update_clicked = st.form_submit_button(
                "Update Bill",
                type="primary",
                use_container_width=True,
            )

        if update_clicked:
            if not edit_bill_no.strip():
                st.error("Bill No. is required.")
            else:
                duplicate = any(
                    r["_excel_row"] != row_num
                    and str(r["Bill No."] or "").strip().casefold()
                    == edit_bill_no.strip().casefold()
                    for r in load_records()
                )

                if duplicate:
                    st.error(
                        "Another record already uses this Bill No."
                    )
                else:
                    updated = make_record(
                        edit_date,
                        edit_bill_no,
                        numeric_value(edit_rupees),
                        edit_party,
                        numeric_value(edit_without),
                        numeric_value(edit_cgst),
                        numeric_value(edit_sgst),
                        numeric_value(edit_total),
                    )

                    try:
                        update_record(row_num, updated)
                        st.success("Bill updated successfully.")
                        st.rerun()
                    except PermissionError:
                        st.error(
                            "Close the Excel file and try again."
                        )

        st.divider()
        st.warning("Deleting a bill removes it from this workbook.")

        confirm_delete = st.checkbox(
            f"Confirm deletion of Bill No. {selected['Bill No.']}",
            key=f"confirm_delete_{row_num}",
        )

        if st.button(
            "🗑️ Delete Selected Bill",
            disabled=not confirm_delete,
        ):
            try:
                delete_record(row_num)
                st.success("Bill deleted successfully.")
                st.rerun()
            except PermissionError:
                st.error("Close the Excel file and try again.")


# ==================================================
# 9. DOWNLOAD EXCEL
# ==================================================

with tab_export:
    st.subheader("Download Monthly Excel File")

    st.write(
        f"Download the records stored in **{EXCEL_FILE.name}**."
    )

    try:
        excel_bytes = make_excel_download()

        st.download_button(
            "📥 Download Excel File",
            data=excel_bytes,
            file_name=EXCEL_FILE.name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
    except PermissionError:
        st.error("Close the Excel file and try again.")
