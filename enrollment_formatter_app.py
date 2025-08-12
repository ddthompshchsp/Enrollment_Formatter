import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image
from datetime import datetime, date

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
    ws.freeze_panes = "B4"  # keep PID column + header frozen

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

        # Fixed title without center name
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
        filter_row = 3
        ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(ws.max_column)}{ws.max_row}"

        for cell in ws[filter_row]:
            cell.font = Font(bold=True)

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for col in range(13, 16):  # M to O
            ws.cell(row=filter_row, column=col).fill = yellow_fill

        red_font = Font(color="FF0000", bold=True)
        cutoff_date = datetime(2025, 5, 11)

        # helper to parse possible date strings without extra libs
        def try_parse_date(v):
            if isinstance(v, datetime):
                return v
            if isinstance(v, date):
                return datetime(v.year, v.month, v.day)
            if isinstance(v, str):
                s = v.strip()
                for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
                    try:
                        return datetime.strptime(s, fmt)
                    except Exception:
                        continue
            # Let openpyxl-marked dates pass through via cell.is_date (handled below)
            return None

        # Missing values -> "X" in red
        # Any date (real Excel date or parsed string) before 5/11/2025 -> "X" in red
        for row in ws.iter_rows(min_row=filter_row + 1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                val = cell.value
                if val in (None, "", "nan"):
                    cell.value = "X"
                    cell.font = red_font
                    continue

                # Check Excel-native date
                is_early = False
                if getattr(cell, "is_date", False) and isinstance(val, (datetime, date)):
                    dt = val if isinstance(val, datetime) else datetime(val.year, val.month, val.day)
                    is_early = dt < cutoff_date
                else:
                    # Try parsing strings like "04/25/2024", "2025-05-01", etc.
                    dt = try_parse_date(val)
                    is_early = dt is not None and dt < cutoff_date

                if is_early:
                    cell.value = "X"
                    cell.font = red_font

        # Final output
        final_output = "Formatted_Enrollment_Checklist.xlsx"
        wb.save(final_output)

        with open(final_output, "rb") as f:
            st.download_button("📥 Download Formatted Excel", f, file_name=final_output)




