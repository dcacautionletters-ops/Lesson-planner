import streamlit as st
import pandas as pd
import io

def process_data(planner_df, attendance_df):
    """
    Consolidates the Lesson Planner and Attendance data.
    """
    # 1. Standardize column names (stripping whitespace)
    planner_df.columns = planner_df.columns.str.strip()
    attendance_df.columns = attendance_df.columns.str.strip()

    # 2. Process Attendance: Get Max Hours Conducted per Batch/Course
    # We group by Batch and Course Name to get the max hours and keep the staff name
    attendance_grouped = attendance_df.groupby(['Batch', 'Course Name']).agg({
        'Hours Conducted': 'max',
        'Staff Name': 'first'  # Keeps the first staff name found for this group
    }).reset_index()

    # 3. Merge: Left join planner data with attendance
    # We use both Batch and Course Name as the unique key
    consolidated = pd.merge(
        planner_df, 
        attendance_grouped, 
        on=['Batch', 'Course Name'], 
        how='left'
    )

    # 4. Final Cleanup: Select and rename columns as requested
    # Mapping to your specific order
    final_df = consolidated[[
        'Batch', 'Course Name', 'Faculty Name', 
        'Session planned', 'No Of Sessions Taken', 
        'Syllabus Coverage (%)', 'Hours Conducted'
    ]]
    
    # Add an index column
    final_df.insert(0, 'Sl No', range(1, len(final_df) + 1))
    
    return final_df

# --- Streamlit UI ---
st.set_page_config(page_title="Lesson Planner Tracker", layout="wide")
st.title("📊 Lesson Planner & Attendance Tracker")

col1, col2 = st.columns(2)

with col1:
    planner_file = st.file_uploader("Upload Lesson Planner Report", type=['xlsx', 'csv'])
with col2:
    attendance_file = st.file_uploader("Upload Attendance Report", type=['xlsx', 'csv'])

if planner_file and attendance_file:
    # Read files
    # Note: Using header search logic to handle the row 3/4 issue
    planner_df = pd.read_csv(planner_file, skiprows=5) # Adjust skiprows based on file structure
    
    # Logic for attendance start row
    # We read the first 5 rows to identify which row the "Batch" header starts on
    temp_df = pd.read_csv(attendance_file, nrows=5)
    # Automatically finding the row where 'Batch' exists
    skip_rows = 0
    for i in range(len(temp_df)):
        if "Batch" in temp_df.iloc[i].values:
            skip_rows = i
            break
            
    attendance_df = pd.read_csv(attendance_file, skiprows=skip_rows)

    if st.button("Process & Consolidate"):
        with st.spinner("Processing data..."):
            result_df = process_data(planner_df, attendance_df)
            
            st.success("Data consolidated successfully!")
            st.dataframe(result_df)
            
            # Download button
            csv = result_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                "Download Consolidated Report", 
                data=csv, 
                file_name="Consolidated_Report.csv",
                mime="text/csv"
            )
else:
    st.info("Please upload both files to proceed.")
