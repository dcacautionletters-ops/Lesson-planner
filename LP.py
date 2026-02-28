import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- 1. DATA PARSING ENGINE ---
def parse_hours_excel(file):
    """Parses the Hours Conducted Excel file (handles multi-block data)."""
    # Read the whole file without a set header
    df = pd.read_excel(file, header=None)
    data = []
    
    current_batch = None
    for i, row in df.iterrows():
        line = str(row[0]).strip()
        
        # Identify the batch block
        if "SUMMARY" in line.upper():
            current_batch = line.split(" -")[0].strip()
        
        # Identify the header row for the data block
        elif "SUBJECT NAME" in line.upper():
            # Read the next rows as data until we hit a blank row or next batch
            header_row = row
            for j in range(i + 1, len(df)):
                row_data = df.iloc[j]
                if pd.isna(row_data[0]) or "SUMMARY" in str(row_data[0]).upper():
                    break
                
                # Logic to capture Subject and Hours
                # Adjust column indices based on: Subject (0), StaffA (1), HRS (2), StaffB (3), HRS (4)...
                subject = row_data[0]
                # Loop through columns looking for HRS
                for k in range(2, len(row_data), 2):
                    if pd.notna(row_data[k]) and isinstance(row_data[k], (int, float)):
                        data.append({'Batch': current_batch, 'Subject': subject, 'Hours': row_data[k]})
    return pd.DataFrame(data)

# --- 2. STYLING ENGINE ---
def apply_pro_styling(filename):
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
        while j <= ws.max_row and ws.cell(row=j, column=2).value == ws.cell(row=i, column=2).value:
            j += 1
        if j - i > 1:
            ws.merge_cells(start_row=i, start_column=2, end_row=j-1, end_column=2)
            ws.cell(row=i, column=2).alignment = Alignment(vertical='center', horizontal='center')
        
        # Border logic
        for row in range(i, j):
            for col in range(1, ws.max_column + 1):
                ws.cell(row=row, column=col).border = thin_border
        i = j
    wb.save(filename)

# --- 3. APP UI ---
st.set_page_config(page_title="Pro Report Generator", layout="wide")
st.title("🎓 Professional WD Report Generator")

# Inputs (Accepting only xlsx)
lp_file = st.file_uploader("Upload Lesson Planner (.xlsx)", type=['xlsx'])
hc_files = st.file_uploader("Upload Hours Conducted Files (.xlsx)", type=['xlsx'], accept_multiple_files=True)

if lp_file and hc_files and st.button("Generate Professional Report"):
    try:
        # Load Planner
        df_lp = pd.read_excel(lp_file, header=5)
        
        # Parse Hours
        all_hours_df = pd.concat([parse_hours_excel(f) for f in hc_files])
        
        # Merge Logic (Simplified)
        # You would perform your pd.merge here based on Subject Name and Batch
        
        # Sort and Format
        batch_order = ['BCA 2025', 'BCA 2024', 'BCA 2023']
        df_lp['Sort_Key'] = pd.Categorical(df_lp.iloc[:, 2], categories=batch_order, ordered=True)
        df_final = df_lp.sort_values(['Course Name', 'Sort_Key'])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name="Professional Report")
        
        apply_pro_styling(output)
        
        st.success("Report Generated Successfully!")
        st.download_button("Download Report", data=output.getvalue(), file_name="Pro_Report.xlsx", 
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    except Exception as e:
        st.error(f"Error: {e}")
