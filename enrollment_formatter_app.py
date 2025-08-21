# app.py
from datetime import datetime, date
from zoneinfo import ZoneInfo

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
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    cutoff_date = datetime(2025, 5, 11)

    # Helpers
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

    def most_recent(series):
        dates, texts = [], []
        for v in pd.unique(series.dropna()):
            dt = coerce_to_dt(v)
            if dt:
                dates.append(dt)
            else:
                texts.append(v)
        if dates:
            return max(dates)
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
    # 3) Write temp workbook
    # ----------------------------
    title = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    temp_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=0)
        pd.DataFrame([[timestamp]]).to_excel(writer, index=False, header=False, startrow=1)
        df.to_excel(writer, index=False, startrow=3)

    # ----------------------------
    # 4) Style with openpyxl
    # ----------------------------
    wb = load_workbook(temp_path)
    ws = wb.active

    filter_row = 4
    data_start = filter_row + 1
    data_end = ws.max_row
    max_col = ws.max_column

    ws.freeze_panes = "B4"
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(max_col)}{data_end}"

    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"].font = Font(size=10, italic=True, color="555555")

    # Header styling (blue + wrap text)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    red_font = Font(color="FF0000", bold=True)

    # Check/format data cells
    for r in range(data_start, data_end + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            cell.border = thin_border

            if val in (None, "", "nan", "NaT"):
                cell.value = "X"
                cell.font = red_font
                continue

            dt = coerce_to_dt(val)
            if dt:
                if dt < cutoff_date:
                    cell.value = "X"
                    cell.font = red_font
                else:
                    cell.value = dt
                    cell.number_format = "m/d/yy"
                continue

    # Set column widths
    width_map = {1: 16, 2: 22}
    for c in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # ----------------------------
    # 5) Totals (one row only)
    # ----------------------------
    headers = [ws.cell(row=filter_row, column=c).value for c in range(1, max_col + 1)]
    name_col_idx = None
    for idx, h in enumerate(headers, start=1):
        if isinstance(h, str) and "name" in h.lower():
            name_col_idx = idx
            break
    if name_col_idx is None:
        name_col_idx = 2

    total_row = ws.max_row + 2

    ws.cell(row=total_row, column=1, value="Total…………").font = Font(bold=True)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

    center = Alignment(horizontal="center", vertical="center")
    top_border = Border(top=Side(style="thin"))

    total_cells = max(0, data_end - data_start + 1)

    for c in range(1, max_col + 1):
        cell = ws.cell(row=total_row, column=c)
        if c <= name_col_idx:
            cell.value = None
            continue

        valid_count = 0
        for r in range(data_start, data_end + 1):
            if ws.cell(row=r, column=c).value != "X":
                valid_count += 1

        cell.value = valid_count
        cell.alignment = center
        cell.font = Font(bold=True)
        cell.border = top_border

    # ----------------------------
    # 6) Save and download
    # ----------------------------
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)

    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)


