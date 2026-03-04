import streamlit as st
import pandas as pd
import io

# Use the robust loader above
def load_and_fix_headers(uploaded_file):
    uploaded_file.seek(0)
    df_preview = pd.read_csv(uploaded_file, header=None, nrows=20, encoding='latin1')
    header_idx = 0
    for idx, row in df_preview.iterrows():
        if row.astype(str).str.contains('Batch').any():
            header_idx = idx
            break
    uploaded_file.seek(0)
    df = pd.read_csv(uploaded_file, skiprows=header_idx, encoding='latin1')
    df.columns = df.columns.astype(str).str.strip()
    return df

st.set_page_config(page_title="Lesson Tracker", layout="wide")
st.title("📊 Lesson Planner & Attendance Tracker")

col1, col2 = st.columns(2)
with col1:
    planner_file = st.file_uploader("Upload Lesson Planner Report", type=['csv'])
with col2:
    attendance_file = st.file_uploader("Upload Attendance Report", type=['csv'])

if planner_file and attendance_file:
    if st.button("Generate Consolidated Report"):
        # Load
        df_planner = load_and_fix_headers(planner_file)
        df_attendance = load_and_fix_headers(attendance_file)
        
        # DEBUG: Check if 'Batch' is now in columns
        if 'Batch' not in df_planner.columns:
            st.error(f"Batch column still not found. Planner columns: {df_planner.columns.tolist()}")
        elif 'Batch' not in df_attendance.columns:
            st.error(f"Batch column still not found. Attendance columns: {df_attendance.columns.tolist()}")
        else:
            # Aggregate Attendance
            # Note: Ensure the attendance file has a column named exactly 'Hours Conducted'
            try:
                attendance_grouped = df_attendance.groupby(['Batch', 'Course Name'])['Hours Conducted'].max().reset_index()
                
                # Merge
                final_df = pd.merge(df_planner, attendance_grouped, on=['Batch', 'Course Name'], how='left')
                
                # Column Cleanup (Adjust these to match the exact names found in your debug step)
                final_df = final_df[[
                    'Batch', 'Course Name', 'Faculty Name', 
                    'Session planned', 'No Of Sessions Taken', 
                    'Syllabus Coverage (%)', 'Hours Conducted'
                ]]
                
                st.success("Consolidated successfully!")
                st.dataframe(final_df)
            except Exception as e:
                st.error(f"Processing error: {e}")
