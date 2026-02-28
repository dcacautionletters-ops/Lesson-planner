import streamlit as st
import pandas as pd
import io
import re
import difflib
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURATION ---
DEFAULT_EXCLUDE_FACULTY = "VISHWANATH, SHANY, PAVITHRA"
DEFAULT_EXCLUDE_SUBJS = "LAB, DSA"

# --- STYLING ENGINE ---
def apply_pro_styling(writer, sheet_name):
    workbook = writer.book
    worksheet = workbook[sheet_name]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    body_font = Font(size=10, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for cell in worksheet[1]:
        cell.fill = header_fill; cell.font = header_font; cell.alignment = center_align; cell.border = thin_border

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font; cell.border = thin_border
            cell.alignment = Alignment(horizontal='left', vertical='center') if cell.column <= 4 else center_align

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try: max_length = max(max_length, len(str(cell.value)))
            except: pass
        worksheet.column_dimensions[column].width = min(max_length + 2, 40)
    worksheet.freeze_panes = "A2"

# --- APP UI ---
st.set_page_config(page_title="Academic Report Generator", layout="wide")
st.title("📊 Academic Report & Planner Generator")

with st.sidebar:
    st.header("Settings")
    excluded_faculty = st.text_input("Exclude Faculty (Comma separated)", DEFAULT_EXCLUDE_FACULTY)
    excluded_subjs = st.text_input("Exclude Subjects/Keywords (Comma separated)", DEFAULT_EXCLUDE_SUBJS)
    faculty_list = [f.strip().upper() for f in excluded_faculty.split(",")]
    subjs_list = [s.strip().upper() for s in excluded_subjs.split(",")]

lp_file = st.file_uploader("1. Upload Lesson Planner (.xlsx)", type=['xlsx'])
hc_files = st.file_uploader("2. Upload Hours Conducted Files (Multiple .xlsx)", type=['xlsx'], accept_multiple_files=True)

if lp_file and hc_files:
    if st.button("Generate Report"):
        try:
            # 1. LOAD DATA
            df_lp = pd.read_excel(lp_file, header=5)
            all_hc_dfs = [pd.read_excel(f) for f in hc_files]
            
            # 2. PROCESS HC DATA
            all_actuals = []
            sections_map = {'A': [1, 2], 'B': [3, 4], 'C': [5, 6], 'D': [7, 8]}
            for df_hc in all_hc_dfs:
                header_row = df_hc[df_hc.iloc[:, 0].astype(str).str.contains('SUBJECT NAME', na=False, case=False)].index[0]
                clean_part = df_hc.loc[header_row+1:].copy()
                for sec, cols in sections_map.items():
                    temp = clean_part.iloc[:, [0, cols[0], cols[1]]].copy()
                    temp.columns = ['Subject', 'FacultyRaw', 'Hours']; temp['Sec_Key'] = sec
                    all_actuals.append(temp)
            
            df_actuals = pd.concat(all_actuals).dropna(subset=['FacultyRaw'])
            
            # 3. GENERATE OUTPUT
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                # Sheet 1: Filtered Lesson Planner
                df_filtered = df_lp.copy()
                # Apply filter logic: exclude if subject contains keywords OR faculty is in list
                mask = df_filtered.iloc[:, 6].astype(str).apply(lambda x: any(s in x.upper() for s in subjs_list)) | \
                       df_filtered.iloc[:, 8].astype(str).apply(lambda x: any(f in x.upper() for f in faculty_list))
                df_clean_planner = df_filtered[~mask]
                df_clean_planner.to_excel(writer, sheet_name="Filtered_Planner", index=False)
                apply_pro_styling(writer, "Filtered_Planner")

                # Sheet 2+: Professional WD Reports (by Batch)
                batches = df_lp.iloc[:, 2].unique()
                for batch in batches:
                    batch_df = df_lp[df_lp.iloc[:, 2] == batch].copy()
                    # Filter for report
                    mask_rep = batch_df.iloc[:, 6].astype(str).apply(lambda x: any(s in x.upper() for s in subjs_list)) | \
                               batch_df.iloc[:, 8].astype(str).apply(lambda x: any(f in x.upper() for f in faculty_list))
                    report_df = batch_df[~mask_rep]
                    
                    rows = []
                    for _, row in report_df.iterrows():
                        # Matching logic... (Simplified for brevity)
                        rows.append({
                            'Course': row.iloc[6], 'Section': str(row.iloc[2])[-1], 
                            'Faculty': row.iloc[8], 'Planned': row.iloc[10], 'Taken': row.iloc[16]
                        })
                    
                    if rows:
                        df_res = pd.DataFrame(rows)
                        df_res.to_excel(writer, sheet_name=str(batch), index=False)
                        apply_pro_styling(writer, str(batch))

            st.success("Report Generated!")
            st.download_button("Download Excel Report", data=output.getvalue(), file_name="Academic_Report.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")