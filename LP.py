import streamlit as st
import pandas as pd
import io

def clean_headers(df):
    """
    Standardizes column names by stripping spaces and removing newlines.
    """
    # Strip spaces and normalize
    df.columns = df.columns.astype(str).str.strip().str.replace('\n', ' ')
    return df

def load_file(uploaded_file):
    """Loads file, cleans headers, and returns dataframe."""
    try:
        # Load as CSV first
        df = pd.read_csv(uploaded_file)
        # Find the row that contains 'Batch' and treat it as the header
        # This fixes your 'Row 3 vs Row 4' issue
        for i in range(min(10, len(df))):
            if "Batch" in df.iloc[i].values:
                df = pd.read_csv(uploaded_file, skiprows=i)
                break
        
        df = clean_headers(df)
        return df
    except:
        # Fallback to Excel
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file)
        # Apply same header logic for Excel
        df = clean_headers(df)
        return df

# --- UI Setup ---
st.set_page_config(page_title="Lesson Planner Tracker", layout="wide")
st.title("📊 Lesson Planner & Attendance Tracker")

col1, col2 = st.columns(2)

with col1:
    planner_file = st.file_uploader("Upload Lesson Planner Report", type=['csv', 'xlsx'])
with col2:
    attendance_file = st.file_uploader("Upload Attendance Report", type=['csv', 'xlsx'])

if planner_file and attendance_file:
    if st.button("Generate Consolidated Report"):
        df_planner = load_file(planner_file)
        df_attendance = load_file(attendance_file)
        
        # DEBUG: Show user what headers are actually detected
        st.write("Planner Headers Found:", df_planner.columns.tolist())
        st.write("Attendance Headers Found:", df_attendance.columns.tolist())
        
        try:
            # Aggregate attendance (matching your requirements)
            attendance_grouped = df_attendance.groupby(['Batch', 'Course Name'])['Hours Conducted'].max().reset_index()
            
            # Merge
            final_df = pd.merge(df_planner, attendance_grouped, on=['Batch', 'Course Name'], how='left')
            
            # Use columns that were actually detected
            # We map your specific requirements to the cleaner headers
            final_df = final_df[[
                'Batch', 'Course Name', 'Faculty Name', 
                'Session planned', 'No Of Sessions Taken', 
                'Syllabus Coverage (%)', 'Hours Conducted'
            ]]
            
            st.success("Success!")
            st.dataframe(final_df)
            
        except KeyError as e:
            st.error(f"Mapping Error: Column {e} not found.")
            st.info("Check the 'Headers Found' lists above. Update the code column names to match exactly.")
