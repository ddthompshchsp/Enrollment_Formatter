# app.py
import io
from datetime import datetime, date

import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

# ----------------------------
# Streamlit setup (must be first)
# ----------------------------
st.set_page_config(page_title="Enrollment Formatter", layout="centered")

# ----------------------------
# Header / UI
# ----------------------------
try:
    logo = Image.open("header_logo.png")  # file should live next to app.py
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Checklist Formatter (2025–2026)")
st.markdown("Upload your **Enrollment.xlsx** file to receive a formatted version.")

uploaded_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"])

if uploaded_file:
    # ----------------------------
    # 1) Find the header row in the source
    # ----------------------------
    # We first read with openpyxl to detect the row containing "ST: Participant PID"
    wb_src = load_workbook(uploaded_file, data_only=True)
    ws_src = wb_src.active

    header_row = None
    for row in ws_src.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if isinstance(cell.value, str) and "ST: Participant PID" in cell.value:
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        st.error("Couldn't find 'ST: Participant PID' in the file. Please upload the correct file.")
        st.stop()

    # IMPORTANT: rewind the file pointer before pandas reads it again
    uploaded_file.seek(0)

    # ----------------------------
    # 2) Load the table into pandas from the detected header row
    # ----------------------------
    df = pd.read_excel(uploaded_file, header=header_row - 1)
    # Normalize column names by removing "ST: "
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    # Date cutoff: anything before this is considered "early" -> mark "X"
    cutoff_date = datetime(2025, 5, 11)

    # Helper to coerce a value to datetime if possible
    def coerce_to_dt(v):
        if pd.isna(v):
            return None
        if isinstance(v, datetime):
            return v
        if isinstance(v, date):
            return datetime(v.year, v.month, v.day)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            try:
                return from_excel(v)
            except Exception:
                return None
        if isinstance(v, str):
            for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(v.strip(), fmt)
                except Exception:
                    continue
        return None

    # Collapse rows so each PID has one row, keeping the most recent valid date per column
    def most_recent(series):
        # Collect all datetimes we can parse
        dates = []
        texts = []
        for v in pd.unique(series.dropna()):
            dt = coerce_to_dt(v)
            if dt:
                dates.append(dt)
            else:
                texts.append(v)
        if dates:
            return max(dates)
        # fallback: first non-null text/value
        return texts[0] if texts else None

    if "Participant PID" not in df.columns:
        st.error("The file is missing the 'Participant PID' column after parsing.")
        st.stop()

    df = (
        df.dropna(subset=["Participant PID"])
          .groupby("Participant PID", as_index=False)
          .agg(most_recent)
    )

    # ----------------------------
    # 3) Write a temporary workbook (pandas), then re-open to style (openpyxl)
    # ----------------------------
    title = "Enrollment Checklist 2025–2026"
    timestamp = datetime.now().strftime("Generated on %B %d, %Y at %I:%M %p")

    temp_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        # Title rows
        pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=0)
        pd.DataFrame([[timestamp]]).to_excel(writer, index=False, header=False, startrow=1)
        # Data starts on row 4 (1-indexed) -> startrow=3 (0-indexed)
        df.to_excel(writer, index=False, startrow=3)

    # ----------------------------
    # 4) Style with openpyxl
    # ----------------------------
    wb = load_workbook(temp_path)
    ws = wb.active

    filter_row = 4  # header row in the output file
    data_start = filter_row + 1
    data_end = ws.max_row
    max_col = ws.max_column

    # Freeze panes to keep PID column + header visible
    ws.freeze_panes = "B4"

    # AutoFilter
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    # Title styling
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"].font = Font(size=10, italic=True, color="555555")

    # Header styling
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill

    # Optional: highlight specific columns (M–O) if they exist
    yellow_fill = PatternFill(start_color="FFF7AE", end_color="FFF7AE", fill_type="solid")
    for col in range(13, min(16, max_col + 1)):  # 13=M, 14=N, 15=O
        ws.cell(row=filter_row, column=col).fill = yellow_fill

    # Border template
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Fonts for missing/early
    red_font = Font(color="FF0000", bold=True)

    # Validate cells: mark missing/early dates as "X", otherwise leave as-is
    # Also convert any valid dates to real Excel dates with the format m/d/yy
    for r in range(data_start, data_end + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value

            # Apply a border to all data cells for a clean grid look
            cell.border = thin_border

            # Treat empty strings / NaN-like as missing
            if val in (None, "", "nan", "NaT"):
                cell.value = "X"
                cell.font = red_font
                continue

            # If it's a date, check cutoff and apply number format
            dt = coerce_to_dt(val)
            if dt:
                if dt < cutoff_date:
                    cell.value = "X"
                    cell.font = red_font
                else:
                    cell.value = dt  # ensure it's a datetime
                    cell.number_format = "m/d/yy"
                continue

            # Non-date values: leave them alone

    # Set column widths (PID and a few others wider; others auto-ish)
    width_map = {1: 16, 2: 22}  # A, B
    for c in range(1, max_col + 1):
        col_letter = get_column_letter(c)
        ws.column_dimensions[col_letter].width = width_map.get(c, 14)

    # ----------------------------
    # 5) Totals at the bottom
    # ----------------------------
    first_total_row = ws.max_row + 2
    valid_row = first_total_row
    missing_row = first_total_row + 1

    ws.cell(row=valid_row, column=1, value="Total ✔ (valid)").font = Font(bold=True)
    ws.cell(row=missing_row, column=1, value="Total X (missing/early)").font = Font(bold=True)

    center = Alignment(horizontal="center", vertical="center")
    top_border = Border(top=Side(style="thin"))

    for c in range(1, max_col + 1):
        missing_count = 0
        total_cells = max(0, data_end - data_start + 1)

        for r in range(data_start, data_end + 1):
            if ws.cell(row=r, column=c).value == "X":
                missing_count += 1

        valid_count = max(total_cells - missing_count, 0)

        vcell = ws.cell(row=valid_row, column=c, value=valid_count)
        mcell = ws.cell(row=missing_row, column=c, value=missing_count)

        vcell.alignment = center
        mcell.alignment = center
        vcell.font = Font(bold=True)
        mcell.font = Font(bold=True)
        vcell.border = top_border
        mcell.border = top_border

    # Left-align the labels in column A
    ws.cell(row=valid_row, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=missing_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

    # ----------------------------
    # 6) Save and offer download
    # ----------------------------
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)

    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)





