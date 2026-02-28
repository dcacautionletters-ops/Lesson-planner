import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. STYLING ENGINE ---
def format_excel_file(filename):
    wb = load_workbook(filename)
    ws = wb.active
    
    # Define styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Style Header
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font; cell.alignment = center; cell.border = thin_border

    # Merge Subject cells (Column B is index 2)
    i = 2
    while i <= ws.max_row:
        j = i + 1
        # Check if subject in row J is same as row I
        while j <= ws.max_row and ws.cell(row=j, column=2).value == ws.cell(row=i, column=2).value:
            j += 1
        
        # If consecutive rows found, merge
        if j - i > 1:
            ws.merge_cells(start_row=i, start_column=2, end_row=j-1, end_column=2)
            ws.cell(row=i, column=2).alignment = Alignment(vertical='center', horizontal='center')
        
        # Apply borders to the range
        for row in range(i, j):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).border = thin_border
        i = j
    
    wb.save(filename)

# --- 2. APP UI ---
st.set_page_config(page_title="Pro Report Generator", layout="wide")
st.title("🎓 Professional WD Report Generator")

with st.sidebar:
    st.header("Blacklist Configuration")
    ex_fac = st.text_input("Exclude Faculty", "VISHWANATH, SHANY, PAVITHRA")
    ex_sub = st.text_input("Exclude Keywords", "LAB, DSA")
    faculty_list = [f.strip().upper() for f in ex_fac.split(",")]
    subjs_list = [s.strip().upper() for s in ex_sub.split(",")]

lp_file = st.file_uploader("1. Upload Lesson Planner (.xlsx)", type=['xlsx'])
hc_files = st.file_uploader("2. Upload Hours Conducted Files", type=['xlsx'], accept_multiple_files=True)

if lp_file and hc_files:
    if st.button("Generate Pro Report"):
        try:
            # Load Data
            df_lp = pd.read_excel(lp_file, header=5)
            
            # --- FILTERING LOGIC ---
            # Exclude rows containing blacklisted faculty or subjects
            mask = df_lp.iloc[:, 6].astype(str).apply(lambda x: any(s in x.upper() for s in subjs_list)) | \
                   df_lp.iloc[:, 8].astype(str).apply(lambda x: any(f in x.upper() for f in faculty_list))
            
            df_final = df_lp[~mask].copy()
            
            # Sorting logic: BCA 2025 -> 2024 -> 2023
            batch_order = ['BCA 2025', 'BCA 2024', 'BCA 2023']
            df_final['Sort_Key'] = pd.Categorical(df_final.iloc[:, 2], categories=batch_order, ordered=True)
            df_final = df_final.sort_values(['Course Name', 'Sort_Key'])
            
            # --- EXCEL CREATION ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name="Professional Report", index=False)
            
            # --- APPLY FORMATTING ---
            format_excel_file(output)
            
            st.success("Report Generated!")
            st.download_button("Download Professional Report", data=output.getvalue(), 
                               file_name="Pro_Academic_Report.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        except Exception as e:
            st.error(f"Error: {e}")
