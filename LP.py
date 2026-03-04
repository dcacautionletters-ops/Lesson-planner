import streamlit as st
import pandas as pd
import io

# --- Helper Functions ---

def find_header_row(file_obj, search_term="Batch"):
    """Scans the file to find the row index where the header starts."""
    # We read a preview of the file to find the header row
    try:
        temp_df = pd.read_csv(file_obj, header=None, nrows=20, encoding='latin1')
        file_obj.seek(0) # Reset file pointer after reading
        for index, row in temp_df.iterrows():
            if search_term in row.values:
                return index
    except Exception:
        return 0
    return 0

def load_data(uploaded_file):
    """Loads CSV/Excel files robustly."""
    header_idx = find_header_row(uploaded_file)
    
    # Try loading with utf-8, fallback to latin1
    try:
        df = pd.read_csv(uploaded_file, skiprows=header_idx, encoding='utf-8')
    except:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, skiprows=header_idx, encoding='latin1')
        
    df.columns = df.columns.str.strip() # Remove extra spaces from headers
    return df

# --- Main App ---

st.set_page_config(page_title="Lesson Planner Tracker", layout="wide")
st.title("📊 Lesson Planner & Attendance Tracker")

st.markdown("""
### Instructions:
1. Upload your **Lesson Planner Report** and **Attendance Report**.
2. The system will automatically detect the data headers.
3. Click **'Generate Consolidated Report'** to merge and process the files.
""")

col1, col2 = st.columns(2)

with col1:
    planner_file = st.file_uploader("Upload Lesson Planner Report", type=['csv', 'xlsx'])
with col2:
    attendance_file = st.file_uploader("Upload Attendance Report", type=['csv', 'xlsx'])

if planner_file and attendance_file:
    if st.button("Generate Consolidated Report"):
        with st.spinner("Processing data..."):
            # Load Data
            planner_df = load_data(planner_file)
            attendance_df = load_data(attendance_file)
            
            # 1. Process Attendance: Get Max Hours per Batch & Course
            # We group to ensure we capture the maximum hours conducted if duplicates exist
            attendance_grouped = attendance_df.groupby(['Batch', 'Course Name']).agg({
                'Hours Conducted': 'max'
            }).reset_index()
            
            # 2. Merge Data
            # Merging on Batch and Course Name
            final_df = pd.merge(
                planner_df, 
                attendance_grouped, 
                on=['Batch', 'Course Name'], 
                how='left'
            )
            
            # 3. Rename and Reorder Columns to your requirement
            # Mapping existing columns to the final format
            try:
                # Selecting and cleaning the output
                output_df = final_df[[
                    'Batch', 'Course Name', 'Faculty Name', 
                    'Session planned', 'No Of Sessions Taken', 
                    'Syllabus Coverage (%)', 'Hours Conducted'
                ]].copy()
                
                # Add Serial Number
                output_df.insert(0, 'Sl No', range(1, len(output_df) + 1))
                
                st.success("Successfully Consolidated!")
                st.dataframe(output_df, use_container_width=True)
                
                # Download
                csv_buffer = io.BytesIO()
                output_df.to_csv(csv_buffer, index=False)
                st.download_button(
                    label="Download Excel/CSV Report",
                    data=csv_buffer.getvalue(),
                    file_name="Consolidated_Lesson_Report.csv",
                    mime="text/csv"
                )
            except KeyError as e:
                st.error(f"Column Mapping Error: Please check if the column names match exactly. Missing: {e}")
                st.write("Current columns found:", final_df.columns.tolist())
