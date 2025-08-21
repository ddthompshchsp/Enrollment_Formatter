import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from PIL import Image
from datetime import datetime, date

# Streamlit UI
logo = Image.open("header_logo.png")  # Must be in the same directory
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

        cutoff_date = datetime(2025, 5, 11)

        # ✅ Collapse rows so each PID has one row, keeping the most recent valid date per column
        def most_recent(series):
            # try converting everything to datetimes
            dates = []
            for v in series.dropna().unique():
                if isinstance(v, (datetime, date)):
                    dt = v if isinstance(v, datetime) else datetime(v.year, v.month, v.day)
                    dates.append(dt)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    try:
                        dt = from_excel(v)
                        dates.append(dt)
                    except Exception:
                        continue
                elif isinstance(v, str):
                    for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
                        try:
                            dt = datetime.strptime(v.strip(), fmt)
                            dates.append(dt)
                            break
                        except Exception:
                            continue
            if dates:
                return max(dates)  # return the most recent date
            else:
                # if not a date, return first non-null value (text, status, etc.)
                vals = series.dropna().unique()
                return vals[0] if len(vals) > 0 else None

        df = (
            df.dropna(subset=["Participant PID"])
              .groupby("Participant PID", as_index=False)
              .agg(most_recent)
        )

        # Title & date
        title = "Enrollment Checklist 2025–2026"
        timestamp = datetime.now().strftime("Generated on %B %d, %Y at %I:%M %p")

        # Save interim file
        temp_path = "Enrollment_Cleaned.xlsx"
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            pd.DataFrame([[title]]).to_excel(writer, index=False, header=False, startrow=0)
            pd.DataFrame([[timestamp]]).to_excel(writer, index=False, header=False, startrow=1)
            df.to_excel(writer, index=False, startrow=3)

        # Step 3: Load for formatting
        wb = load_workbook(temp_path)
        ws = wb.active

        # Formatting
        filter_row = 4
        ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(ws.max_column)}{ws.max_row}"

        # Title styling
        ws["A1"].font = Font(size=14, bold=True)
        ws["A2"].font = Font(size=10, italic=True, color="555555")

        # Header styling
        for cell in ws[filter_row]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Highlight specific columns (M to O, adjust if needed)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for col in range(13, 16):
            ws.cell(row=filter_row, column=col).fill = yellow_fill

        # Fonts for missing/early
        red_font = Font(color="FF0000", bold=True)

        # ✅ Check all cells for missing/early dates
        for row in ws.iter_rows(min_row=filter_row + 1, max_row=ws.max_row,
                                min_col=1, max_col=ws.max_column):
            for cell in row:
                val = cell.value

                if val in (None, "", "nan"):
                    cell.value = "X"
                    cell.font = red_font
                    continue

                if isinstance(val, (datetime, date)):
                    dt = val if isinstance(val, datetime) else datetime(val.year, val.month, val.day)
                    if dt < cutoff_date:
                        cell.value = "X"
                        cell.font = red_font
                    continue

        # Final output
        final_output = "Formatted_Enrollment_Checklist.xlsx"
        wb.save(final_output)

        with open(final_output, "rb") as f:
            st.download_button("📥 Download Formatted Excel", f, file_name=final_output)




