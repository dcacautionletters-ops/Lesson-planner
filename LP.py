import streamlit as st
import pandas as pd
import io
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. DATA PARSING ENGINE ---
def parse_hours_file(file):
    """Reads the nested summary file and flattens it into a clean table."""
    df = pd.read_csv(file, header=None)
    data = []
    
    batch = None
    for i, row in df.iterrows():
        line = str(row[0]).strip()
        if "SUMMARY" in line:
            batch = line.split(" -")[0].strip()
        elif "SUBJECT NAME" in line:
            # We found the header, read the rows until empty or next batch
            for j in range(i + 1, len(df)):
                row_data = df.iloc[j]
                if pd.isna(row_data[0]) or "SUMMARY" in str(row_data[0]): break
                # Extract subject and hours (Assuming pattern: Sub, StaffA, HrsA, StaffB, HrsB...)
                subject = row_data[0]
                # Logic to grab hours for each section (A, B, C, D...)
                # Adjust column indices based on your specific CSV structure
                for k in range(1, len(row_data), 2):
                    if pd.notna(row_data[k]) and k+1 < len(row_data):
                        data.append({'Batch': batch, 'Subject': subject, 'Section': chr(64 + (k+1)//2), 'Hours': row_data[k+1]})
    return pd.DataFrame(data)

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

    # Merge Subject cells (Column B is index 2)
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
st.set_page_config(page_title="Professional Report Generator", layout="wide")
st.title("🎓 Professional WD Report Generator")

# Inputs
lp_file = st.file_uploader("Upload Lesson Planner (.xlsx)", type=['xlsx'])
hc_files = st.file_uploader("Upload Hours Conducted (.csv)", type=['csv'], accept_multiple_files=True)

if lp_file and hc_files and st.button("Generate Report"):
    try:
        # Load Planner
        # Automatically find the header row
        temp_df = pd.read_excel(lp_file, header=None)
        header_idx = temp_df[temp_df.apply(lambda row: row.astype(str).str.contains("Course Name").any(), axis=1)].index[0]
        df_lp = pd.read_excel(lp_file, header=header_idx)
        df_lp.columns = df_lp.columns.str.strip()

        # Parse Hours
        all_hours = pd.concat([parse_hours_file(f) for f in hc_files])
        
        # Merge/Process
        # [Merging Logic Here...]
        
        # Sort and Format
        batch_order = ['BCA 2025', 'BCA 2024', 'BCA 2023']
        df_lp['Sort_Key'] = pd.Categorical(df_lp['Batch'], categories=batch_order, ordered=True)
        df_final = df_lp.sort_values(['Course Name', 'Sort_Key'])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Professional Report")
        
        apply_pro_styling(output)
        st.download_button("Download Report", data=output.getvalue(), file_name="Pro_Report.xlsx")
        st.success("Report generated successfully!")
        
    except Exception as e:
        st.error(f"Error: {e}")
