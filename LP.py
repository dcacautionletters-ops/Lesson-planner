import streamlit as st
import pandas as pd
import difflib
import re
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURATION ---
EXCLUDE_FACULTY = ["VISHWANATH", "SHANY", "PAVITHRA"]

def get_closest_match(name, possibilities, cutoff=0.6):
    """Fuzzy match for faculty names with lower cutoff for better matching."""
    if not name or pd.isna(name): return None
    name_clean = str(name).upper().strip()
    matches = difflib.get_close_matches(name_clean, possibilities, n=1, cutoff=cutoff)
    return matches[0] if matches else None

def apply_pro_styling(writer, sheet_name):
    """Applies corporate styling to the Excel sheet"""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    for col in worksheet.columns:
        worksheet.column_dimensions[col[0].column_letter].width = 22

def process_attendance(uploaded_file):
    """Processes the attendance report based on the uploaded file structure."""
    if uploaded_file is None: return pd.DataFrame()
    
    # Based on uploaded file: Row 3 is header (Index 2)
    df = pd.read_excel(uploaded_file, header=2)
    
    # Mapping: G=Batch(6), I=Course(8), J=Hours Conducted(9), Q=Staff Name(16)
    try:
        df_clean = pd.DataFrame({
            'Batch': df.iloc[:, 6],
            'Subject': df.iloc[:, 8],
            'Hours': pd.to_numeric(df.iloc[:, 9], errors='coerce'),
            'FacultyRaw': df.iloc[:, 16]
        }).dropna(subset=['FacultyRaw'])
    except:
        return pd.DataFrame()
    
    processed = []
    for _, row in df_clean.iterrows():
        names = re.split(r',| AND ', str(row['FacultyRaw']), flags=re.IGNORECASE)
        section = str(row['Batch']).strip()[-1] if len(str(row['Batch'])) > 0 else ""
        for n in names:
            processed.append({
                'Subject': str(row['Subject']).upper().strip(),
                'Faculty': n.upper().strip(),
                'Hours': row['Hours'] or 0,
                'Sec_Key': section
            })
    return pd.DataFrame(processed)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Academic Report Builder", layout="wide")
st.title("📊 Final Academic Report Generator")

col1, col2, col3 = st.columns(3)
with col1: lp_file = st.file_uploader("1. Raw Lesson Planner", type=['xlsx'])
with col2: mca_hc_file = st.file_uploader("2. Attendance MCA", type=['xlsx'])
with col3: bca_hc_file = st.file_uploader("3. Attendance BCA", type=['xlsx'])

if all([lp_file, mca_hc_file, bca_hc_file]):
    if st.button("Generate Final Consolidated Report"):
        try:
            # 1. Load Lesson Planner (Header is usually row 6 / Index 5)
            df_lp = pd.read_excel(lp_file, header=5)
            
            # 2. Combine Attendance Data
            att_data = pd.concat([process_attendance(mca_hc_file), process_attendance(bca_hc_file)], ignore_index=True)
            
            batch_keys = ["BCA 2023", "BCA 2024", "BCA 2025", "MCA 2024", "MCA 2025"]
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for is_lab in [False, True]:
                    label = "LAB_REPORTS" if is_lab else "THEORY_REPORTS"
                    final_rows = []

                    for key in batch_keys:
                        # Filter Planner for specific Batch
                        batch_df = df_lp[df_lp.iloc[:, 2].astype(str).str.contains(key, na=False)].copy()
                        
                        for _, lp_row in batch_df.iterrows():
                            course = str(lp_row.iloc[6]).upper().strip()
                            faculty = str(lp_row.iloc[8]).upper().strip()
                            
                            if any(ex in faculty for ex in EXCLUDE_FACULTY): continue
                            if (is_lab and "LAB" not in course) or (not is_lab and "LAB" in course): continue

                            batch_full = str(lp_row.iloc[2]).strip()
                            section = batch_full[-1]
                            
                            # Matching
                            matched_staff = get_closest_match(faculty, att_data['Faculty'].unique().tolist())
                            sub_match = difflib.get_close_matches(course, att_data['Subject'].unique(), n=1, cutoff=0.5)

                            actual_hrs_from_log = 0
                            if sub_match and matched_staff:
                                match_filter = att_data[
                                    (att_data['Sec_Key'] == section) & 
                                    (att_data['Subject'] == sub_match[0]) &
                                    (att_data['Faculty'] == matched_staff)
                                ]
                                actual_hrs_from_log = match_filter['Hours'].max() if not match_filter.empty else 0

                            final_rows.append({
                                'Batch': batch_full,
                                'Course Name': lp_row.iloc[6],
                                'Faculty Name': lp_row.iloc[8],
                                'Planned Sessions': pd.to_numeric(lp_row.iloc[10], errors='coerce') or 0,
                                'Sessions Taken (LP)': lp_row.iloc[16],
                                'Syllabus %': lp_row.iloc[18],
                                'Actual Hours (Attendance Log)': actual_hrs_from_log
                            })

                    if final_rows:
                        df_sheet = pd.DataFrame(final_rows).groupby(['Course Name', 'Batch'], as_index=False).agg({
                            'Faculty Name': lambda x: ', '.join(x.unique()),
                            'Planned Sessions': 'max',
                            'Sessions Taken (LP)': 'max',
                            'Syllabus %': 'max',
                            'Actual Hours (Attendance Log)': 'max'
                        })
                        df_sheet.insert(0, 'Sl No.', range(1, len(df_sheet) + 1))
                        df_sheet.to_excel(writer, sheet_name=label, index=False)
                        apply_pro_styling(writer, label)

            st.success("✅ Report Ready!")
            st.download_button("📥 Download Excel", output.getvalue(), "Academic_Consolidated_Report.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")
