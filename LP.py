import streamlit as st
import pandas as pd
import difflib
import re
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURATION ---
EXCLUDE_FACULTY = ["VISHWANATH", "SHANY", "PAVITHRA"]

def get_closest_match(name, possibilities, cutoff=0.70):
    if not name or pd.isna(name): return None
    name_clean = str(name).upper().strip()
    matches = difflib.get_close_matches(name_clean, possibilities, n=1, cutoff=cutoff)
    return matches[0] if matches else None

def apply_pro_styling(writer, sheet_name):
    workbook = writer.book
    if sheet_name not in workbook.sheetnames: return
    worksheet = workbook[sheet_name]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name='Calibri')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    for col in worksheet.columns:
        worksheet.column_dimensions[col[0].column_letter].width = 25

def extract_section(batch_name):
    batch_str = str(batch_name).strip()
    if len(batch_str) > 2 and batch_str[-2] == " " and batch_str[-1].isalpha():
        return batch_str[-1]
    return batch_str

def process_attendance_file(uploaded_file):
    """Processes any Attendance Report (BCA, MCA, or Consolidated)."""
    if uploaded_file is None: return pd.DataFrame()
    try:
        df = pd.read_excel(uploaded_file, header=2)
        # Mapping: Col 6: Batch, Col 8: Course, Col 9: Hours, Col 16: Staff
        df_clean = pd.DataFrame({
            'Batch': df.iloc[:, 6],
            'Subject': df.iloc[:, 8],
            'Hours': pd.to_numeric(df.iloc[:, 9], errors='coerce'),
            'FacultyRaw': df.iloc[:, 16]
        }).dropna(subset=['FacultyRaw'])
        
        processed = []
        for _, row in df_clean.iterrows():
            names = re.split(r',| AND ', str(row['FacultyRaw']), flags=re.IGNORECASE)
            sec_key = extract_section(row['Batch'])
            for n in names:
                processed.append({
                    'Subject': str(row['Subject']).upper().strip(),
                    'Faculty': n.upper().strip(),
                    'Hours': row['Hours'] or 0,
                    'Sec_Key': sec_key
                })
        return pd.DataFrame(processed).groupby(['Subject', 'Faculty', 'Sec_Key'], as_index=False).max()
    except Exception:
        return pd.DataFrame()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Universal Academic Master", layout="wide")
st.title("📑 Academic Reporting Hub")
st.info("Flexibility: Upload attendance files separately OR as one single consolidated file. The system will process whatever is provided.")

# Layout for Uploads
lp_col, att_col = st.columns([1, 2])

with lp_col:
    st.subheader("1. Lesson Planner")
    lp_file = st.file_uploader("Upload Master Lesson Planner", type=['xlsx'], key="lp")

with att_col:
    st.subheader("2. Attendance Data")
    # Multiple files can be uploaded to this single box, or used across multiple boxes
    att_files = st.file_uploader("Upload Attendance Files (BCA / MCA / Consolidated)", 
                                 type=['xlsx'], 
                                 accept_multiple_files=True, 
                                 key="att_files")

if lp_file and att_files:
    if st.button("🚀 Generate Consolidated Report"):
        try:
            # 1. Load Planner
            df_lp = pd.read_excel(lp_file, header=5)
            
            # 2. Process and Merge all uploaded Attendance files
            all_att_dfs = [process_attendance_file(f) for f in att_files]
            att_data = pd.concat(all_att_dfs, ignore_index=True) if all_att_dfs else pd.DataFrame()
            
            if att_data.empty:
                st.error("No valid attendance data found in uploaded files.")
                st.stop()

            # Universal Batch Keys
            batch_keys = ["BCA 2023", "BCA 2024", "BCA 2025", "BCA AIML", "BCA DS", "MCA 2024", "MCA 2025"]
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sheets_created = 0
                for is_lab in [False, True]:
                    label = "LAB_SESSIONS" if is_lab else "THEORY_SESSIONS"
                    final_rows = []

                    for key in batch_keys:
                        batch_df = df_lp[df_lp.iloc[:, 2].astype(str).str.contains(key, na=False)].copy()
                        
                        for _, lp_row in batch_df.iterrows():
                            course = str(lp_row.iloc[6]).upper().strip()
                            faculty = str(lp_row.iloc[8]).upper().strip()
                            
                            if any(ex in faculty for ex in EXCLUDE_FACULTY): continue
                            if (is_lab and "LAB" not in course) or (not is_lab and "LAB" in course): continue

                            batch_full = str(lp_row.iloc[2]).strip()
                            sec_key = extract_section(batch_full)
                            
                            m_faculty = get_closest_match(faculty, att_data['Faculty'].unique().tolist())
                            m_course = difflib.get_close_matches(course, att_data['Subject'].unique(), n=1, cutoff=0.5)

                            actual_hrs = 0
                            if m_course and m_faculty:
                                match = att_data[(att_data['Sec_Key'] == sec_key) & 
                                                 (att_data['Subject'] == m_course[0]) &
                                                 (att_data['Faculty'] == m_faculty)]
                                actual_hrs = match['Hours'].iloc[0] if not match.empty else 0

                            final_rows.append({
                                'Batch': batch_full,
                                'Course Name': lp_row.iloc[6],
                                'Faculty Name': lp_row.iloc[8],
                                'Planned Sessions': pd.to_numeric(lp_row.iloc[10], errors='coerce') or 0,
                                'As per Time Table': lp_row.iloc[11],      # Col L
                                'No of sessions taken': lp_row.iloc[16],   # Col Q
                                'Syllabus Coverage %': lp_row.iloc[18],    # Col S
                                'Actual Hours Conducted': actual_hrs
                            })

                    if final_rows:
                        df_final = pd.DataFrame(final_rows).groupby(['Course Name', 'Batch'], as_index=False).agg({
                            'Faculty Name': lambda x: ', '.join(x.unique()),
                            'Planned Sessions': 'max',
                            'As per Time Table': 'max',
                            'No of sessions taken': 'max',
                            'Syllabus Coverage %': 'max',
                            'Actual Hours Conducted': 'max'
                        })
                        
                        # LOGIC: Actual - Planned
                        df_final['Deviation'] = df_final['Actual Hours Conducted'] - df_final['Planned Sessions']
                        df_final.insert(0, 'Sl No.', range(1, len(df_final) + 1))
                        
                        df_final.to_excel(writer, sheet_name=label, index=False)
                        apply_pro_styling(writer, label)
                        sheets_created += 1

                if sheets_created == 0:
                    pd.DataFrame({"Status": ["No matches found"]}).to_excel(writer, sheet_name="Empty")

            st.success("✨ Report Compiled!")
            st.download_button("📥 Download Master Report", output.getvalue(), "Academic_Report.xlsx")

        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.warning("Please upload the Lesson Planner and at least one Attendance file to begin.")
