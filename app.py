
from pathlib import Path
from datetime import date, datetime
import calendar
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------


APP_DIR = Path(__file__).resolve().parent

# Default to the previous month's sales period
today = date.today()

if today.month == 1:
    sales_year = today.year - 1
    sales_month = 12
else:
    sales_year = today.year
    sales_month = today.month - 1

sales_month_name = calendar.month_abbr[sales_month].upper()

default_filename = (
    f"GST SALES {sales_month_name} {sales_year}"
)

st.session_state.setdefault(
    "active_workbook_name",
    default_filename
)


EXCEL_FILE = (
    APP_DIR / f"{st.session_state['active_workbook_name']}.xlsx"
)


COLUMNS = [
    "Date",
    "Bill No.",
    "Rupees (₹)",
    "GST Credit Party Name",
    "Without GST (₹)",
    "9%",
    "9%",
    "With Final Total (₹)",
]

st.set_page_config(
    page_title="GST Bill Management System",
    page_icon="🧾",
    layout="wide",
)


# --------------------------------------------------
# EXCEL FUNCTIONS
# --------------------------------------------------

def style_worksheet(ws):
    """Apply professional formatting to the Excel sheet."""

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

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

    for row in range(2, ws.max_row + 1):
        for col in (3, 5, 6, 7, 8):
            ws.cell(row, col).number_format = '#,##0.00'

        ws.cell(row, 1).number_format = "dd-mm-yyyy"


def ensure_workbook():
    """Create the Excel file if it does not exist."""

    if EXCEL_FILE.exists():
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Bill Records"
    ws.append(COLUMNS)

    style_worksheet(ws)
    wb.save(EXCEL_FILE)


def load_records():
    """Read saved bills and retain their Excel row numbers."""

    ensure_workbook()

    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb["Bill Records"]

    records = []

    for excel_row in range(2, ws.max_row + 1):
        values = [
            ws.cell(excel_row, col).value
            for col in range(1, len(COLUMNS) + 1)
        ]

        if all(value is None for value in values):
            continue

        record = dict(zip(COLUMNS, values))
        record["_excel_row"] = excel_row
        records.append(record)

    wb.close()
    return records


def save_record(record):
    """Append one new bill to the Excel workbook."""

    ensure_workbook()

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    ws.append([record[column] for column in COLUMNS])

    style_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def update_record(excel_row, record):
    """Update an existing bill without adding a duplicate row."""

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    for col, name in enumerate(COLUMNS, start=1):
        ws.cell(excel_row, col).value = record[name]

    style_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def delete_record(excel_row):
    """Delete an existing bill."""

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    ws.delete_rows(excel_row, 1)

    style_worksheet(ws)
    wb.save(EXCEL_FILE)
    wb.close()


def make_excel_download():
    """Create an Excel download from the current saved file."""

    ensure_workbook()

    wb = load_workbook(EXCEL_FILE)
    ws = wb["Bill Records"]

    style_worksheet(ws)

    output = BytesIO()
    wb.save(output)
    wb.close()

    output.seek(0)
    return output.getvalue()


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
    """Build a record using the required Excel column names."""

    return {
        "Date": datetime.combine(bill_date, datetime.min.time()),
        "Bill No.": bill_no.strip(),
        "Rupees (₹)": float(rupees),
        "GST Credit Party Name": party.strip(),
        "Without GST (₹)": float(without_gst),
        "9% - CGST (₹)": float(cgst),
        "9% - SGST (₹)": float(sgst),
        "With Final Total (₹)": float(final_total),
    }


def validate_amounts(*amounts):
    """Reject negative or invalid monetary values."""

    for amount in amounts:
        if pd.isna(amount) or float(amount) < 0:
            return False

    return True


# --------------------------------------------------
# APPLICATION HEADER
# --------------------------------------------------

st.title("🧾 GST Bill Management System")
st.caption(
    "Enter, save, search, edit and manage bill records "
    "directly in an Excel workbook."
)

ensure_workbook()
records = load_records()

total_bills = len(records)
total_value = sum(
    float(record["With Final Total (₹)"] or 0)
    for record in records
)

col1, col2, col3 = st.columns(3)

col1.metric("Total Bills", total_bills)
col2.metric("Total Final Amount", f"₹ {total_value:,.2f}")
col3.metric(
    "Excel File",
    "Available",
    help=str(EXCEL_FILE),
)

st.divider()


# --------------------------------------------------
# NAVIGATION
# --------------------------------------------------

tab_add, tab_view, tab_edit, tab_export = st.tabs(
    [
        "➕ Add Bill",
        "📋 View & Search",
        "✏️ Edit / Delete",
        "📥 Excel Download",
    ]
)


# --------------------------------------------------
# ADD BILL
# --------------------------------------------------

with tab_add:
    st.subheader("Enter New Bill")

    with st.form("add_bill_form", clear_on_submit=True):

        col1, col2 = st.columns(2)

        with col1:
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
                value=0.0,
                step=100.0,
                format="%.2f",
                help="Independent numeric amount.",
            )

            party = st.text_input(
                "4. GST Credit Party Name",
                placeholder="Enter party name",
            )

        with col2:
                without_gst = st.number_input(
                "5. Without GST (₹)",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
                )

                cgst = st.number_input(
                    "6. 9% - CGST (₹)",
                    min_value=0.0,
                    value=0.0,
                    step=10.0,
                    format="%.2f",
                )

                sgst = st.number_input(
                    "7. 9% - SGST (₹)",
                    min_value=0.0,
                    value=0.0,
                    step=10.0,
                    format="%.2f",
                )


               calculated_total = without_gst + cgst + sgst
        
            if auto_total:
                st.info(
                    "Calculated final total: "
                    f"₹ {calculated_total:,.2f}"
                )
                final_total = calculated_total
            else:
                final_total = st.number_input(
                    "8. With Final Total (₹)",
                    min_value=0.0,
                    value=float(calculated_total),
                    step=100.0,
                    format="%.2f",
                )

        submitted = st.form_submit_button(
            "💾 Save Bill",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        if not bill_no.strip():
            st.error("Please enter the Bill No.")

        elif not party.strip():
            st.error("Please enter the GST Credit Party Name.")

        elif not validate_amounts(
            rupees,
            without_gst,
            cgst,
            sgst,
            final_total,
        ):
            st.error("All amounts must be valid and non-negative.")

        else:
            duplicate = any(
                str(r["Bill No."]).strip().casefold()
                == bill_no.strip().casefold()
                for r in load_records()
            )

            if duplicate:
                st.error(
                    "This Bill No. already exists. "
                    "Use Edit / Delete to update the existing bill."
                )
            else:
                record = make_record(
                    bill_date,
                    bill_no,
                    rupees,
                    party,
                    without_gst,
                    cgst,
                    sgst,
                    final_total,
                )

                try:
                    save_record(record)
                    st.success(
                        f"Bill {bill_no.strip()} saved successfully."
                    )
                    st.rerun()

                except PermissionError:
                    st.error(
                        "Cannot write to bills.xlsx. "
                        "Close the workbook in Excel and try again."
                    )


# --------------------------------------------------
# VIEW AND SEARCH
# --------------------------------------------------

with tab_view:
    st.subheader("Saved Bill Records")

    search = st.text_input(
        "Search by Bill No. or Party Name",
        placeholder="Type a bill number or party name",
        key="search_records",
    )

    current_records = load_records()

    if search.strip():
        term = search.strip().casefold()

        current_records = [
            r for r in current_records
            if term in str(r["Bill No."]).casefold()
            or term in str(r["GST Credit Party Name"]).casefold()
        ]

    if current_records:
        df = pd.DataFrame([
            {column: record[column] for column in COLUMNS}
            for record in current_records
        ])

        df["Date"] = pd.to_datetime(
            df["Date"], errors="coerce"
        ).dt.strftime("%d-%m-%Y")

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )

        st.caption(f"Showing {len(df)} bill(s).")
    else:
        st.info("No matching bills found.")


# --------------------------------------------------
# EDIT / DELETE
# --------------------------------------------------

with tab_edit:
    st.subheader("Edit or Delete a Bill")

    current_records = load_records()

    if not current_records:
        st.info("No bills available to edit.")
    else:
        def display_label(record):
            return (
                f"Row {record['_excel_row']} | "
                f"{record['Bill No.']} | "
                f"{record['GST Credit Party Name']}"
            )

        selected_label = st.selectbox(
            "Select a bill",
            options=[display_label(r) for r in current_records],
        )

        selected = next(
            r for r in current_records
            if display_label(r) == selected_label
        )

        excel_row = selected["_excel_row"]

        stored_date = selected["Date"]
        if isinstance(stored_date, datetime):
            stored_date = stored_date.date()
        elif not isinstance(stored_date, date):
            stored_date = date.today()

        stored_values = [
            float(selected[column] or 0)
            for column in [
                "Rupees (₹)",
                "Without GST (₹)",
                "9% - CGST (₹)",
                "9% - SGST (₹)",
                "With Final Total (₹)",
            ]
        ]

        with st.form("edit_bill_form"):
            st.caption("Update the selected record.")

            edit_date = st.date_input(
                "Date",
                value=stored_date,
                key="edit_date",
            )

            edit_bill_no = st.text_input(
                "Bill No.",
                value=str(selected["Bill No."]),
                key="edit_bill_no",
            )

            edit_rupees = st.number_input(
                "Rupees (₹)",
                min_value=0.0,
                value=stored_values[0],
                step=100.0,
                format="%.2f",
                key="edit_rupees",
            )

            edit_party = st.text_input(
                "GST Credit Party Name",
                value=str(selected["GST Credit Party Name"]),
                key="edit_party",
            )

            edit_without = st.number_input(
                "Without GST (₹)",
                min_value=0.0,
                value=stored_values[1],
                step=100.0,
                format="%.2f",
                key="edit_without",
            )

            edit_cgst = st.number_input(
                "9% - CGST (₹)",
                min_value=0.0,
                value=stored_values[2],
                step=10.0,
                format="%.2f",
                key="edit_cgst",
            )

            edit_sgst = st.number_input(
                "9% - SGST (₹)",
                min_value=0.0,
                value=stored_values[3],
                step=10.0,
                format="%.2f",
                key="edit_sgst",
            )

            edit_auto = st.checkbox(
                "Calculate final total automatically",
                value=True,
                key="edit_auto",
            )

            edit_calculated = edit_without + edit_cgst + edit_sgst

            if edit_auto:
                st.info(
                    f"Final total: ₹ {edit_calculated:,.2f}"
                )
                edit_final = edit_calculated
            else:
                edit_final = st.number_input(
                    "With Final Total (₹)",
                    min_value=0.0,
                    value=stored_values[4],
                    step=100.0,
                    format="%.2f",
                    key="edit_final",
                )

            update_clicked = st.form_submit_button(
                "Update Bill",
                type="primary",
                use_container_width=True,
            )

        if update_clicked:
            if not edit_bill_no.strip() or not edit_party.strip():
                st.error("Bill No. and Party Name are required.")
            elif not validate_amounts(
                edit_rupees,
                edit_without,
                edit_cgst,
                edit_sgst,
                edit_final,
            ):
                st.error("Amounts must be valid and non-negative.")
            else:
                duplicate = any(
                    r["_excel_row"] != excel_row
                    and str(r["Bill No."]).strip().casefold()
                    == edit_bill_no.strip().casefold()
                    for r in load_records()
                )

                if duplicate:
                    st.error(
                        "Another bill already uses this Bill No."
                    )
                else:
                    updated = make_record(
                        edit_date,
                        edit_bill_no,
                        edit_rupees,
                        edit_party,
                        edit_without,
                        edit_cgst,
                        edit_sgst,
                        edit_final,
                    )

                    try:
                        update_record(excel_row, updated)
                        st.success("Bill updated successfully.")
                        st.rerun()
                    except PermissionError:
                        st.error(
                            "Close bills.xlsx in Excel and try again."
                        )

        st.divider()
        st.warning(
            "Deleting a bill permanently removes its row "
            "from the Excel workbook."
        )

        confirm_delete = st.checkbox(
            f"I confirm deletion of bill {selected['Bill No.']}",
            key=f"confirm_delete_{excel_row}",
        )

        if st.button(
            "🗑️ Delete Selected Bill",
            type="secondary",
            disabled=not confirm_delete,
        ):
            try:
                delete_record(excel_row)
                st.success("Bill deleted successfully.")
                st.rerun()
            except PermissionError:
                st.error(
                    "Close bills.xlsx in Excel and try again."
                )


# --------------------------------------------------
# EXCEL DOWNLOAD
# --------------------------------------------------

with tab_export:
    st.subheader("Download Excel Workbook")

    st.write(
        "Download a copy of all saved bills, including "
        "Excel formatting and column headings."
    )

    try:
        excel_bytes = make_excel_download()

        st.download_button(
            label="📥 Download bills.xlsx",
            data=excel_bytes,
            file_name="bills.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )

        st.caption(f"Local workbook: {EXCEL_FILE}")

    except PermissionError:
        st.error("Close the Excel workbook and try again.")
