import streamlit as st
import pandas as pd
import io

# --- Robust File Loader ---
def load_file(uploaded_file):
    """
    Tries to load an Excel file; if it fails, falls back to CSV.
    This solves the 'UnicodeDecodeError' and 'ParserError'.
    """
    # Try reading as Excel first
    try:
        # We skip rows until we find the header
        # Using header=None and finding the row with 'Batch' is safest
        df = pd.read_excel(uploaded_file, header=None)
        
        # Find the row that actually contains the header "Batch"
        header_row = df[df.apply(lambda row: row.astype(str).str.contains('Batch').any(), axis=1)].index[0]
        
        # Re-read with correct header row
        df = pd.read_excel(uploaded_file, header=header_row)
        
    except Exception:
        # Fallback: Treat as CSV
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, header=0) # You may need header=None + logic here if CSVs have meta-rows
    
    # Clean whitespace from column names to prevent merge errors
    df.columns = df.columns.astype(str).str.strip()
    return df

# --- UI Setup ---
st.set_page_config(page_title="Lesson Planner Tracker", layout="wide")
st.title("📊 Lesson Planner & Attendance Tracker")

col1, col2 = st.columns(2)

with col1:
    planner_file = st.file_uploader("Upload Lesson Planner Report (.xlsx/.csv)", type=['xlsx', 'csv'])
with col2:
    attendance_file = st.file_uploader("Upload Attendance Report (.xlsx/.csv)", type=['xlsx', 'csv'])

if planner_file and attendance_file:
    if st.button("Generate Consolidated Report"):
        with st.spinner("Processing files..."):
            try:
                # Load the files
                df_planner = load_file(planner_file)
                df_attendance = load_file(attendance_file)
                
                # Group Attendance to get max hours per batch/course
                # Note: Adjust these column names if your Excel headers are slightly different
                attendance_grouped = df_attendance.groupby(['Batch', 'Course Name'])['Hours Conducted'].max().reset_index()
                
                # Merge
                final_df = pd.merge(
                    df_planner, 
                    attendance_grouped, 
                    on=['Batch', 'Course Name'], 
                    how='left'
                )
                
                # Selection & Reordering
                # Make sure these names match exactly what is in your Excel file
                output_df = final_df[[
                    'Batch', 'Course Name', 'Faculty Name', 
                    'Session planned', 'No Of Sessions Taken', 
                    'Syllabus Coverage (%)', 'Hours Conducted'
                ]].copy()
                
                # Add Serial Number
                output_df.insert(0, 'Sl No', range(1, len(output_df) + 1))
                
                # Show results
                st.success("Successfully Consolidated!")
                st.dataframe(output_df)
                
                # Download
                buffer = io.BytesIO()
                output_df.to_excel(buffer, index=False)
                st.download_button(
                    label="Download Final Excel Report",
                    data=buffer.getvalue(),
                    file_name="Consolidated_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Error processing files: {e}")
                st.write("Tip: Ensure your Excel file headers (Batch, Course Name, etc.) are typed exactly as they appear in the file.")
