import streamlit as st
import pandas as pd
import difflib
import re
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# --- CONFIGURATION ---
EXCLUDE_FACULTY = ["VISHWANATH", "SHANY", "PAVITHRA"]

def get_closest_match(name, possibilities, cutoff=0.75):
    """Fuzzy match for faculty names to bridge naming variations."""
    if not name or pd.isna(name): return None
    name_clean = str(name).upper().strip()
    matches = difflib.get_close_matches(name_clean, possibilities, n=1, cutoff=cutoff)
    return matches[0] if matches else None

def apply_pro_styling(writer, sheet_name):
    """Applies high-end corporate styling to the Excel sheet (replacing the VBA styling)"""
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
            cell.alignment = center_align if cell.column > 4 else Alignment(horizontal='left')

    for col in worksheet.columns:
        column = col[0].column_letter
        worksheet.column_dimensions[column].width = 20

def process_attendance(uploaded_file):
    """Parses the Consolidated Attendance Report format provided."""
    if uploaded_file is None: return pd.DataFrame()
    
    # Reading from Row 3 (Index 2) as per the provided CSV structure
    df = pd.read_excel(uploaded_file, header=2)
    
    # Mapping based on the indices identified: G=Batch(6), I=Course(8), J=Hours(9), Q=Staff(16)
    try:
        df_clean = pd.DataFrame({
            'Batch': df.iloc[:, 6],
            'Subject': df.iloc[:, 8],
            'Hours': pd.to_numeric(df.iloc[:, 9], errors='coerce'),
            'FacultyRaw': df.iloc[:, 16]
        }).dropna(subset=['FacultyRaw'])
    except Exception as e:
        st.error(f"Column mapping error in attendance file: {e}")
        return pd.DataFrame()
    
    processed_rows = []
    for _, row in df_clean.iterrows():
        # Handle shared subjects (split by 'AND' or ',')
        names = re.split(r',| AND ', str(row['FacultyRaw']), flags=re.IGNORECASE)
        # Extract Section (last character of Batch name)
        section = str(row['Batch']).strip()[-1] if len(str(row['Batch'])) > 0 else ""
        
        for n in names:
            processed_rows.append({
                'Subject': str(row['Subject']).upper().strip(),
                'Faculty': n.upper().strip(),
                'Hours': row['Hours'] or 0,
                'Sec_Key': section
            })
    
    return pd.DataFrame(processed_rows)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Academic Planner Pro", layout="wide")
st.title("📊 Academic Report Integrator (Theory & Lab)")

col1, col2, col3 = st.columns(3)
with col1: lp_file = st.file_uploader("1. Raw Lesson Planner", type=['xlsx'])
with col2: mca_hc_file = st.file_uploader("2. Attendance MCA", type=['xlsx'])
with col3: bca_hc_file = st.file_uploader("3. Attendance BCA", type=['xlsx'])

if all([lp_file, mca_hc_file, bca_hc_file]):
    if st.button("🚀 Generate Final Report"):
        try:
            # 1. Load Data
            df_lp = pd.read_excel(lp_file, header=5)
            att_data = pd.concat([process_attendance(mca_hc_file), process_attendance(bca_hc_file)], ignore_index=True)
            
            batch_keys = ["BCA 2023", "BCA 2024", "BCA 2025", "MCA 2024", "MCA 2025"]
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 2. Split Logic: Theory vs Lab
                for is_lab in [False, True]:
                    sheet_name = "LAB_SESSIONS" if is_lab else "THEORY_SESSIONS"
                    combined_rows = []

                    for key in batch_keys:
                        # Filter Lesson Planner for specific Batch group
                        batch_df = df_lp[df_lp.iloc[:, 2].astype(str).str.contains(key, na=False)].copy()
                        
                        for _, lp_row in batch_df.iterrows():
                            course = str(lp_row.iloc[6]).upper().strip()
                            faculty = str(lp_row.iloc[8]).upper().strip()
                            
                            # Filtering Logic
                            if any(ex in faculty for ex in EXCLUDE_FACULTY): continue
                            if (is_lab and "LAB" not in course) or (not is_lab and "LAB" in course):
                                continue

                            batch_full = str(lp_row.iloc[2]).strip()
                            section = batch_full[-1]
                            
                            # Matching using difflib
                            matched_staff = get_closest_match(faculty, att_data['Faculty'].unique().tolist())
                            sub_match = difflib.get_close_matches(course, att_data['Subject'].unique(), n=1, cutoff=0.6)

                            actual_hrs = 0
                            if sub_match and matched_staff:
                                match_filter = att_data[
                                    (att_data['Sec_Key'] == section) & 
                                    (att_data['Subject'] == sub_match[0]) &
                                    (att_data['Faculty'] == matched_staff)
                                ]
                                actual_hrs = match_filter['Hours'].max() if not match_filter.empty else 0

                            combined_rows.append({
                                'Course Name': lp_row.iloc[6],
                                'Section': section,
                                'Batch': batch_full,
                                'Faculty Name': lp_row.iloc[8],
                                'Planned': pd.to_numeric(lp_row.iloc[10], errors='coerce') or 0,
                                'Taken (LP)': lp_row.iloc[16],
                                'Syllabus %': lp_row.iloc[18],
                                'Actual (Log)': actual_hrs
                            })

                    if combined_rows:
                        # 3. Final Aggregation (Group by Course/Section to handle double faculty entries)
                        df_final = pd.DataFrame(combined_rows).groupby(['Course Name', 'Section'], as_index=False).agg({
                            'Batch': 'first',
                            'Faculty Name': lambda x: ', '.join(x.unique()),
                            'Planned': 'max',
                            'Taken (LP)': 'max',
                            'Syllabus %': 'max',
                            'Actual (Log)': 'max'
                        })
                        
                        df_final['Variance'] = df_final['Actual (Log)'] - df_final['Planned']
                        df_final.insert(0, 'Sl No.', range(1, len(df_final) + 1))
                        
                        # Write to Excel and style
                        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
                        apply_pro_styling(writer, sheet_name)

            st.success("✅ Compilation Complete!")
            st.download_button(
                label="📥 Download Consolidated Report",
                data=output.getvalue(),
                file_name="Final_Academic_Report_2024.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Processing Error: {str(e)}")
