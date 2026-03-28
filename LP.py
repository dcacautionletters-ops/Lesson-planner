import streamlit as st
import pandas as pd
import difflib
import re
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURATION ---
EXCLUDE_FACULTY = ["VISHWANATH", "SHANY", "PAVITHRA"]

def get_closest_match(name, possibilities, cutoff=0.75):
    """Fuzzy matching for faculty names."""
    if not name or pd.isna(name): return None
    name_clean = str(name).upper().strip()
    matches = difflib.get_close_matches(name_clean, possibilities, n=1, cutoff=cutoff)
    return matches[0] if matches else None

def apply_pro_styling(writer, sheet_name):
    """Applies corporate styling to the Excel sheet as per your VBA logic."""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    body_font = Font(size=10, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.border = thin_border
            cell.alignment = center_align if cell.column > 4 else Alignment(horizontal='left', vertical='center')

    for col in worksheet.columns:
        column = col[0].column_letter
        worksheet.column_dimensions[column].width = 22

def process_hc_data(df_hc_list):
    """Extracts hours conducted from multi-section attendance logs."""
    all_actuals_list = []
    sections_map = {'A': [1, 2], 'B': [3, 4], 'C': [5, 6], 'D': [7, 8]}

    for df_hc_raw in df_hc_list:
        # Find the row containing 'SUBJECT NAME'
        header_indices = df_hc_raw[df_hc_raw.iloc[:, 0].astype(str).str.contains('SUBJECT NAME', na=False, case=False)].index.tolist()
        for idx in header_indices:
            clean_part = df_hc_raw.loc[idx+1:].copy()
            for sec, cols in sections_map.items():
                if len(clean_part.columns) > max(cols):
                    temp = clean_part.iloc[:, [0, cols[0], cols[1]]].copy()
                    temp.columns = ['Subject', 'FacultyRaw', 'Hours']
                    temp['Sec_Key'] = sec
                    all_actuals_list.append(temp)
    
    if not all_actuals_list: return pd.DataFrame()
    
    combined = pd.concat(all_actuals_list).dropna(subset=['FacultyRaw'])
    processed = []
    for _, row in combined.iterrows():
        names = re.split(r',| AND ', str(row['FacultyRaw']), flags=re.IGNORECASE)
        for n in names:
            processed.append({
                'Subject': str(row['Subject']).upper().strip(),
                'Faculty': n.upper().strip(),
                'Hours': pd.to_numeric(row['Hours'], errors='coerce') or 0,
                'Sec_Key': row['Sec_Key']
            })
    return pd.DataFrame(processed)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Academic Report Pro", layout="wide")
st.title("📊 Integrated Lesson Planner & Attendance Report")

col1, col2, col3 = st.columns(3)
with col1: lp_file = st.file_uploader("1. Lesson Planner", type=['xlsx'])
with col2: mca_hc_file = st.file_uploader("2. Attendance MCA", type=['xlsx'])
with col3: bca_hc_file = st.file_uploader("3. Attendance BCA", type=['xlsx'])

if all([lp_file, mca_hc_file, bca_hc_file]):
    if st.button("Generate Professional Report"):
        try:
            # Load Data
            df_lp = pd.read_excel(lp_file, header=5)
            df_hc_list = [pd.read_excel(mca_hc_file), pd.read_excel(bca_hc_file)]
            df_actuals = process_hc_data(df_hc_list)
            
            batch_keys = ["BCA 2023", "BCA 2024", "BCA 2025", "MCA 2024", "MCA 2025"]
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Iterate for Theory and Labs
                for is_lab in [False, True]:
                    sheet_label = "LAB_SESSIONS" if is_lab else "THEORY_SESSIONS"
                    sheet_rows = []

                    for key in batch_keys:
                        batch_df = df_lp[df_lp.iloc[:, 2].astype(str).str.contains(key, na=False)].copy()
                        
                        for _, lp_row in batch_df.iterrows():
                            course = str(lp_row.iloc[6]).upper().strip()
                            faculty = str(lp_row.iloc[8]).upper().strip()
                            
                            # Filter based on type
                            if is_lab and "LAB" not in course: continue
                            if not is_lab and "LAB" in course: continue
                            if any(ex in faculty for ex in EXCLUDE_FACULTY): continue

                            batch_full = str(lp_row.iloc[2]).strip()
                            section = batch_full[-1]
                            
                            # Match Logic
                            matched_staff = get_closest_match(faculty, df_actuals['Faculty'].unique().tolist())
                            sub_match = difflib.get_close_matches(course, df_actuals['Subject'].unique(), n=1, cutoff=0.7)

                            actual_hrs = 0
                            if sub_match and matched_staff:
                                m_data = df_actuals[(df_actuals['Sec_Key'] == section) & 
                                                    (df_actuals['Subject'] == sub_match[0]) &
                                                    (df_actuals['Faculty'] == matched_staff)]
                                actual_hrs = m_data['Hours'].max() if not m_data.empty else 0

                            sheet_rows.append({
                                'Course Name': lp_row.iloc[6],
                                'Section': section,
                                'Batch': batch_full,
                                'Faculty Name': lp_row.iloc[8],
                                'Planned Sessions': pd.to_numeric(lp_row.iloc[10], errors='coerce') or 0,
                                'As per Time Table': lp_row.iloc[16],
                                'Syllabus Coverage %': lp_row.iloc[18],
                                'Actual Hours Conducted': actual_hrs
                            })

                    if sheet_rows:
                        df_final = pd.DataFrame(sheet_rows).groupby(['Course Name', 'Section'], as_index=False).agg({
                            'Batch': 'first',
                            'Faculty Name': lambda x: ', '.join(x.unique()),
                            'Planned Sessions': 'max',
                            'As per Time Table': 'max',
                            'Syllabus Coverage %': 'max',
                            'Actual Hours Conducted': 'max'
                        })
                        
                        # Deviation logic: Planned - Actual
                        df_final['Deviation'] = df_final['Actual Hours Conducted'] - df_final['Planned Sessions']
                        
                        # Formatting: Add Serial No and sorting
                        df_final = df_final.sort_values(by=['Batch', 'Course Name'])
                        df_final.insert(0, 'Sl No.', range(1, len(df_final) + 1))
                        
                        df_final.to_excel(writer, sheet_name=sheet_label, index=False)
                        apply_pro_styling(writer, sheet_label)

            st.success("✅ Compilation Successful!")
            st.download_button("📥 Download Final Report", output.getvalue(), "Consolidated_Academic_Report.xlsx")

        except Exception as e:
            st.error(f"Error during processing: {e}")
