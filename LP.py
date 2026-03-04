import streamlit as st
import pandas as pd
import io

# --- Professional Grade File Loader ---
def safe_load_csv(file_obj, keyword='Batch'):
    """
    Reads a CSV, ignores garbage metadata/titles at the top,
    and automatically detects the header row.
    """
    file_obj.seek(0)
    # Read the file line by line to find the header index
    lines = file_obj.readlines()
    header_idx = 0
    for i, line in enumerate(lines):
        line_str = line.decode('latin1', errors='ignore')
        if keyword in line_str:
            header_idx = i
            break
    
    file_obj.seek(0)
    # Load using engine='python' for maximum compatibility and skip bad lines
    df = pd.read_csv(
        file_obj, 
        skiprows=header_idx, 
        encoding='latin1', 
        on_bad_lines='skip', 
        engine='python'
    )
    # Clean whitespace from column names
    df.columns = df.columns.str.strip()
    return df

# --- UI Setup ---
st.set_page_config(page_title="Lesson Tracker", layout="wide")
st.title("📊 Lesson Planner & Attendance Tracker")

col1, col2 = st.columns(2)

with col1:
    planner_file = st.file_uploader("Upload Lesson Planner Report", type=['csv'])
with col2:
    attendance_file = st.file_uploader("Upload Attendance Report", type=['csv'])

if planner_file and attendance_file:
    if st.button("Generate Consolidated Report"):
        with st.spinner("Parsing and merging data..."):
            try:
                # 1. Load Data
                df_planner = safe_load_csv(planner_file, keyword='Batch')
                df_attendance = safe_load_csv(attendance_file, keyword='Batch')

                # 2. Process Attendance (Aggregate max hours)
                # Grouping ensures we get 1 record per Batch/Course
                attendance_grouped = df_attendance.groupby(['Batch', 'Course Name'])['Hours Conducted'].max().reset_index()

                # 3. Merge
                # Left join keeps all planner records even if no attendance is found
                merged_df = pd.merge(
                    df_planner, 
                    attendance_grouped, 
                    on=['Batch', 'Course Name'], 
                    how='left'
                )

                # 4. Final Selection & Formatting
                # Ensure the columns match your specific requirement list
                cols_to_keep = [
                    'Batch', 'Course Name', 'Faculty Name', 
                    'Session planned', 'No Of Sessions Taken', 
                    'Syllabus Coverage (%)', 'Hours Conducted'
                ]
                
                # Check if columns exist to prevent KeyError
                missing_cols = [c for c in cols_to_keep if c not in merged_df.columns]
                if missing_cols:
                    st.error(f"Missing columns in report: {missing_cols}")
                    st.write("Available columns:", merged_df.columns.tolist())
                else:
                    final_df = merged_df[cols_to_keep].copy()
                    
                    # Add Sl No
                    final_df.insert(0, 'Sl No', range(1, len(final_df) + 1))
                    
                    # Show result
                    st.success("Consolidation Complete!")
                    st.dataframe(final_df)
                    
                    # Prepare Download
                    csv_data = final_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "Download Consolidated Report", 
                        data=csv_data, 
                        file_name="Final_Consolidated_Report.csv",
                        mime="text/csv"
                    )

            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
                st.write("Tip: Ensure your file headers match exactly (e.g., 'Course Name' vs 'Course_Name').")
