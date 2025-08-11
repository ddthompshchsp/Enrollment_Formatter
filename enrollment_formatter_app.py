import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image
from datetime import datetime, date  # for date comparison

logo = Image.open("header_logo.png")  # Make sure this image is in the same directory
st.image(logo, width=300)

st.set_page_config(page_title="Enrollment Formatter", layout="centered")

st.title("HCHSP Enrollment Checklist Formatter (2025–2026)")
st.markdown("Upload your **Enrollment.xlsx** file to receive a formatted version.")

uploaded_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"])

if uploaded_file:
    # Step 1: Load the workbook to find the header row
    wb = load_workbook(uploaded_file)
    ws = wb.active
    ws.freeze_panes = "B4"  # keep PID (col A) and header visible

    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if isinstance(cell.value, str) and "ST: Participant PID" in cell.value:
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        st.error("Couldn't find 'ST: Participant PID' in the file. Please upload the correct file.")
    else:
        # Step 2: Load into pandas from header_row
        df = pd.read_excel(uploaded_file, header=header_row - 1)
        df.columns = [col.replace("ST: ", "") if isinstance(col, str) else col for col in df.columns]

        # Remove duplicates by PID
        df = df.dropna(subset=["Participant PID"]).drop_duplicates(subset=["Participant PID"])

        # ---- Rename by position (L, M, R, S) ----
        # A=1 -> index 0, so: L=12->11, M=13->12, R=18->17, S=19->18
        col_names = list(df.columns)
        if len(col_names) >= 12:
            col_names[11] = "Immunization"
        if len(col_names) >= 13:
            col_names[12] = "Returning Child Yes/Si"
        if len(col_names) >= 18:
            col_names[17] = "Food Allergies Yes/No"
        if len(col_names) >= 19:
            col_names[18] = "Nutrition Assessment Date"
        df.columns = col_names

        # Title (no school name)
        title = "Enrollment Checklist 2025–2026"

        # Save interim file
        temp_path = "Enrollment_Cleaned.xlsx"
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=0)
            df.to_excel(writer, index=False, startrow=2)

        # Step 3: Load for formatting
        wb = load_workbook(temp_path)
        ws = wb.active

        # Formatting
        filter_row = 3  # header row in the output file
        ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(ws.max_column)}{ws.max_row}"

        # Bold + wrap headers on row 3
        for cell in ws[filter_row]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrapText=True)

        # Highlight ONLY column M (13) on header row
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        if ws.max_column >= 13:
            ws.cell(row=filter_row, column=13).fill = yellow_fill  # column M

        red_font = Font(color="FF0000", bold=True)
        cutoff_date = datetime(2025, 5, 15)

        # --- helper to parse date-like values without changing your flow ---
        def _to_datetime(v):
            if isinstance(v, datetime):
                return v
            if isinstance(v, date):
                return datetime(v.year, v.month, v.day)
            if isinstance(v, (int, float)):
                # leave numeric alone unless Excel marks it as a date (handled via cell.is_date)
                return None
            if isinstance(v, str):
                s = v.strip()
                for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
                    try:
                        return datetime.strptime(s, fmt)
                    except Exception:
                        pass
            return None

        # Missing values -> "X" in red, and any date (true Excel date OR parsed string) < cutoff -> red font
        for row in ws.iter_rows(min_row=filter_row + 1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                val = cell.value
                if val in [None, "", "nan"]:
                    cell.value = "X"
                    cell.font = red_font
                else:
                    make_red = False
                    if getattr(cell, "is_date", False) and isinstance(val, (datetime, date)):
                        dt = val if isinstance(val, datetime) else datetime(val.year, val.month, val.day)
                        make_red = dt < cutoff_date
                    else:
                        dt = _to_datetime(val)
                        make_red = dt is not None and dt < cutoff_date

                    if make_red:
                        cell.font = red_font

        # Final output
        final_output = "Formatted_Enrollment_Checklist.xlsx"
        wb.save(final_output)

        with open(final_output, "rb") as f:
            st.download_button("📥 Download Formatted Excel", f, file_name=final_output)



