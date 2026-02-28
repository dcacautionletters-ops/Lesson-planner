import streamlit as st
import pandas as pd
import numpy as np
import difflib
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
DEFAULT_EXCLUDE_FACULTY = "VISHWANATH, SHANY, PAVITHRA"
DEFAULT_EXCLUDE_SUBJS = "LAB, DSA"

def apply_final_styling(filename):
    """
    Applies professional styling and Merges the 'Subject' column.
    """
    wb = load_workbook(filename)
    ws = wb.active
    
    # 1. Define Styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # 2. Style Headers
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font; cell.alignment = center; cell.border = thin_border

    # 3. Merge 'Subject' Column (Column B)
    # Start checking from row 2
    i = 2
    while i <= ws.max_row:
        j = i + 1
        while j <= ws.max_row and ws[f'B{j}'].value == ws[f'B{i}'].value:
            j += 1
        # Merge if multiple rows found
        if j - i > 1:
            ws.merge_cells(f'B{i}:B{j-1}')
            ws[f'B{i}'].alignment = Alignment(vertical='center', horizontal='center')
        i = j

    # 4. Global Borders
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            if cell.row > 1: cell.alignment = Alignment(vertical='center')

    wb.save(filename)

# --- APP UI ---
st.set_page_config(page_title="Pro Academic Report", layout="wide")
st.title("🎓 Professional WD Report Generator")

# Sidebar
with st.sidebar:
    st.header("Blacklist Configuration")
    ex_fac = st.text_input("Exclude Faculty", DEFAULT_EXCLUDE_FACULTY)
    ex_sub = st.text_input("Exclude Keywords", DEFAULT_EXCLUDE_SUBJS)

lp_file = st.file_uploader("1. Upload Lesson Planner (.xlsx)", type=['xlsx'])
hc_files = st.file_uploader("2. Upload Hours Conducted Files", type=['xlsx'], accept_multiple_files=True)

if lp_file and hc_files:
    if st.button("Generate Pro Report"):
        try:
            # Load Data
            df_lp = pd.read_excel(lp_file, header=5)
            # Standardize Batch Order (BCA 2025 -> 2024 -> 2023)
            batch_order = ['BCA 2025', 'BCA 2024', 'BCA 2023']
            
            # (Process logic here - simplified for brevity)
            # ... (The data cleaning remains similar to the previous version) ...
            
            # Prepare Data for Output
            # Mapping columns from your snippet:
            # Sl No., Subject, Batch/Section, Faculty, Planned, Taken, Coverage, Actual, Variance
            
            # Create Excel in Memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Group by Subject to ensure merging works
                final_df = df_lp.sort_values(by=['Course Name', 'Batch']) 
                final_df.to_excel(writer, sheet_name="Report", index=False)
                
            # Apply Style & Merge
            apply_final_styling(output)
            
            st.success("Report Generated Successfully!")
            st.download_button("Download Professional Report", data=output.getvalue(), file_name="Pro_Academic_Report.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")
