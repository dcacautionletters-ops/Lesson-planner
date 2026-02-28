import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. DYNAMIC DATA LOADING ---
def load_planner_data(file):
    """Automatically finds the header row and loads the Excel file."""
    # Load first few rows to find where the header is
    temp_df = pd.read_excel(file, header=None)
    
    # Find the row index that contains "Course Name"
    header_idx = temp_df[temp_df.apply(lambda row: row.astype(str).str.contains("Course Name", na=False).any(), axis=1)].index[0]
    
    # Load the actual data using that index
    df = pd.read_excel(file, header=header_idx)
    # Strip whitespace from column names to prevent matching errors
    df.columns = df.columns.str.strip()
    return df

# --- 2. STYLING ENGINE ---
def apply_pro_styling(filename):
    wb = load_workbook(filename)
    ws = wb.active
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Style Header
    for cell in ws[1]:
        cell.fill = header_fill; cell.font = header_font; cell.alignment = center; cell.border = thin_border

    # Merge Subject cells (Column 'Course Name' - usually index 2, checking headers)
    # Assuming 'Course Name' is the 2nd column
    i = 2
    while i <= ws.max_row:
        j = i + 1
        while j <= ws.max_row and ws.cell(row=j, column=2).value == ws.cell(row=i, column=2).value:
            j += 1
        if j - i > 1:
            ws.merge_cells(start_row=i, start_column=2, end_row=j-1, end_column=2)
            ws.cell(row=i, column=2).alignment = Alignment(vertical='center', horizontal='center')
        i = j
    wb.save(filename)

# --- 3. APP UI ---
st.set_page_config(page_title="Pro Report Generator", layout="wide")
st.title("🎓 Professional WD Report Generator")

lp_file = st.file_uploader("Upload Lesson Planner (.xlsx)", type=['xlsx'])
hc_files = st.file_uploader("Upload Hours Conducted Files (.xlsx)", type=['xlsx'], accept_multiple_files=True)

if lp_file and hc_files and st.button("Generate Professional Report"):
    try:
        # Load Planner with automatic header detection
        df_lp = load_planner_data(lp_file)
        
        # Sort logic
        batch_order = ['BCA 2025', 'BCA 2024', 'BCA 2023']
        # Ensure we are checking the right column name
        df_lp['Sort_Key'] = pd.Categorical(df_lp['Batch'], categories=batch_order, ordered=True)
        df_final = df_lp.sort_values(['Course Name', 'Sort_Key'])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Professional Report")
        
        apply_pro_styling(output)
        
        st.success("Report Generated Successfully!")
        st.download_button("Download Report", data=output.getvalue(), file_name="Pro_Academic_Report.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    except Exception as e:
        st.error(f"Error: {e}")
