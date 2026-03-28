import streamlit as st
import pandas as pd
import difflib
import re
import io

# --- CONFIGURATION & UTILS ---
EXCLUDE_FACULTY = ["VISHWANATH", "SHANY", "PAVITHRA"]

def get_closest_match(name, possibilities, cutoff=0.75):
    if not name or pd.isna(name): return None
    name_clean = str(name).upper().strip()
    matches = difflib.get_close_matches(name_clean, possibilities, n=1, cutoff=cutoff)
    return matches[0] if matches else None

def process_attendance(df_hc_list):
    """Processes BCA/MCA Attendance logs into a unified flat format"""
    all_actuals_list = []
    sections_map = {'A': [1, 2], 'B': [3, 4], 'C': [5, 6], 'D': [7, 8]}

    for df_hc_raw in df_hc_list:
        # Find where the data actually starts
        header_indices = df_hc_raw[df_hc_raw.iloc[:, 0].astype(str).str.contains('SUBJECT NAME', na=False, case=False)].index.tolist()
        for idx in header_indices:
            clean_part = df_hc_raw.loc[idx+1:].copy()
            for sec, cols in sections_map.items():
                if len(clean_part.columns) > max(cols):
                    temp = clean_part.iloc[:, [0, cols[0], cols[1]]].copy()
                    temp.columns = ['Subject', 'FacultyRaw', 'Hours']
                    temp['Sec_Key'] = sec
                    all_actuals_list.append(temp)
    
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
st.title("🎓 Professional Academic Report Generator")

col1, col2, col3 = st.columns(3)
with col1: lp_file = st.file_uploader("1. Raw Lesson Planner", type=['xlsx'])
with col2: mca_hc_file = st.file_uploader("2. Attendance MCA", type=['xlsx'])
with col3: bca_hc_file = st.file_uploader("3. Attendance BCA", type=['xlsx'])

if all([lp_file, mca_hc_file, bca_hc_file]):
    if st.button("Generate Consolidated Report"):
        try:
            # Load Data
            df_lp = pd.read_excel(lp_file, header=5)
            df_actuals = process_attendance([pd.read_excel(mca_hc_file), pd.read_excel(bca_hc_file)])
            
            batch_keys = ["BCA 2023", "BCA 2024", "BCA 2025", "MCA 2024", "MCA 2025"]
            
            # Prepare Excel Buffer
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                
                for is_lab in [False, True]:
                    type_label = "LAB" if is_lab else "THEORY"
                    all_type_rows = []

                    for key in batch_keys:
                        batch_df = df_lp[df_lp.iloc[:, 2].astype(str).str.contains(key, na=False)].copy()
                        
                        for _, lp_row in batch_df.iterrows():
                            course = str(lp_row.iloc[6]).upper().strip()
                            faculty = str(lp_row.iloc[8]).upper().strip()
                            
                            # Filter based on Theory vs Lab
                            if any(ex in faculty for ex in EXCLUDE_FACULTY): continue
                            if is_lab and "LAB" not in course: continue
                            if not is_lab and "LAB" in course: continue

                            batch_full = str(lp_row.iloc[2]).strip()
                            section = batch_full[-1] # Grabs A, B, C, or D
                            
                            matched_staff = get_closest_match(faculty, df_actuals['Faculty'].unique().tolist())
                            sub_match = difflib.get_close_matches(course, df_actuals['Subject'].unique(), n=1, cutoff=0.7)

                            actual_hrs = 0
                            if sub_match and matched_staff:
                                m_data = df_actuals[(df_actuals['Sec_Key'] == section) & 
                                                    (df_actuals['Subject'] == sub_match[0]) &
                                                    (df_actuals['Faculty'] == matched_staff)]
                                actual_hrs = m_data['Hours'].max() if not m_data.empty else 0

                            all_type_rows.append({
                                'Batch_Key': key,
                                'Course Name': lp_row.iloc[6],
                                'Section': section,
                                'Batch': batch_full,
                                'Faculty Name': lp_row.iloc[8],
                                'Session Planned': pd.to_numeric(lp_row.iloc[10], errors='coerce') or 0,
                                'Sessions Taken': lp_row.iloc[16],
                                'Syllabus Coverage %': lp_row.iloc[18],
                                'Actual Hours (Log)': actual_hrs
                            })

                    if all_type_rows:
                        df_final = pd.DataFrame(all_type_rows)
                        # Grouping to ensure unique entries per course/section
                        df_final = df_final.groupby(['Batch_Key', 'Course Name', 'Section'], as_index=False).agg({
                            'Batch': 'first',
                            'Faculty Name': lambda x: ', '.join(x.unique()),
                            'Session Planned': 'max',
                            'Sessions Taken': 'max',
                            'Syllabus Coverage %': 'max',
                            'Actual Hours (Log)': 'max'
                        })
                        df_final['Variance'] = df_final['Actual Hours (Log)'] - df_final['Session Planned']
                        df_final.to_excel(writer, sheet_name=type_label, index=False)

            st.success("✅ Report Compiled!")
            st.download_button(
                label="📥 Download Professional Report",
                data=output.getvalue(),
                file_name="Academic_Consolidated_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Logic Error: {e}")
