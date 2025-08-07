
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from openpyxl import load_workbook

st.set_page_config(page_title="Enrollment Formatter", layout="centered")

st.title("📝 Enrollment Checklist Formatter 2025–2026")
st.write("Upload your `Enrollment.xlsx` file below to generate a properly formatted checklist.")

uploaded_file = st.file_uploader("Upload Enrollment.xlsx", type="xlsx")

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)

        # Extract column names from row 4 (index 3)
        column_headers = df_raw.iloc[3].fillna('').astype(str).str.replace("ST: ", "").str.strip()
        df_data = df_raw.iloc[4:].copy()
        df_data.columns = column_headers

        # Remove rows without PID
        df_data = df_data[df_data["Participant PID"].notna()]

        # Trim columns after "Lead Risk Questionnaire: Entered By"
        end_col = "Lead Risk Questionnaire: Entered By"
        if end_col in df_data.columns:
            df_data = df_data.loc[:, :end_col]

        # Consolidate by PID
        df_unique = df_data.groupby("Participant PID", as_index=False).agg(lambda x: x.dropna().iloc[0] if not x.dropna().empty else "")

        # Write to Excel with formatting
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_unique.to_excel(writer, index=False, startrow=1, sheet_name="Checklist")
            wb = writer.book
            ws = writer.sheets["Checklist"]

            # Insert title in row 1
            center_name = df_data["Center Name"].dropna().iloc[0] if "Center Name" in df_data.columns else ""
            title = f"Enrollment Checklist 2025–2026 – {center_name}"
            ws["A1"] = title
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
            ws["A1"].font = Font(bold=True)

            # Bold headers in row 2
            for cell in ws[2]:
                cell.font = Font(bold=True)

            # Highlight columns M-O (columns 13–15)
            yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            for col in range(13, 16):
                ws.cell(row=2, column=col).fill = yellow_fill

            # Add red "X" to missing values from row 3 onward
            red_font = Font(color="FF0000")
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    if cell.value in [None, "", "nan"]:
                        cell.value = "X"
                        cell.font = red_font

            # Set filter starting from row 2
            max_col_letter = get_column_letter(ws.max_column)
            ws.auto_filter.ref = f"A2:{max_col_letter}{ws.max_row}"

        st.success("✅ File formatted successfully! Download below.")
        st.download_button(
            label="📥 Download Formatted Checklist",
            data=output.getvalue(),
            file_name="EnrollmentChecklist_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
