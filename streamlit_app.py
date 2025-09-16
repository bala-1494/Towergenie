import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import os
import time
from datetime import datetime

def create_zip_download(csv_files):
    """Create a ZIP file containing all CSV files for download"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, csv_content in csv_files:
            zip_file.writestr(filename, csv_content)
    zip_buffer.seek(0)
    return zip_buffer

def process_excel_data(df, group_columns, selected_columns, max_rows_per_file, progress_bar=None, status_text=None):
    """Process the dataframe and create CSV files based on grouping criteria"""
    # Filter dataframe to only include selected columns
    df_filtered = df[selected_columns].copy()
    
    # Create groups based on selected columns
    if len(group_columns) == 1:
        grouped = df_filtered.groupby(group_columns[0])
    else:
        grouped = df_filtered.groupby(group_columns)
    
    csv_files = []
    total_groups = 0
    total_files = 0
    
    # Get total number of groups for progress tracking
    total_expected_groups = len(grouped)
    
    # Process each group
    for group_idx, (group_name, group_data) in enumerate(grouped):
        total_groups += 1
        
        # Update progress
        if status_text:
            status_text.text(f"Processing group {total_groups}/{total_expected_groups}: {group_name}")
        if progress_bar:
            progress_bar.progress((group_idx + 1) / total_expected_groups)
        
        # Handle group name formatting
        if isinstance(group_name, tuple):
            group_name_str = "_".join([str(x).replace(" ", "_").replace("/", "_") for x in group_name])
        else:
            group_name_str = str(group_name).replace(" ", "_").replace("/", "_")
        
        # Remove any problematic characters for filename
        group_name_str = "".join(c for c in group_name_str if c.isalnum() or c in ('_', '-'))
        
        # Split group into multiple files if it exceeds max_rows_per_file
        num_rows = len(group_data)
        if num_rows <= max_rows_per_file:
            # Single file for this group
            filename = f"{group_name_str}.csv"
            csv_content = group_data.to_csv(index=False)
            csv_files.append((filename, csv_content))
            total_files += 1
        else:
            # Multiple files for this group
            num_chunks = (num_rows + max_rows_per_file - 1) // max_rows_per_file
            for i in range(num_chunks):
                start_idx = i * max_rows_per_file
                end_idx = min((i + 1) * max_rows_per_file, num_rows)
                chunk_data = group_data.iloc[start_idx:end_idx]
                
                filename = f"{group_name_str}_part_{i+1}.csv"
                csv_content = chunk_data.to_csv(index=False)
                csv_files.append((filename, csv_content))
                total_files += 1
        
        # Add 5-second delay after processing each group (for Streamlit Cloud timeout handling)
        if status_text:
            status_text.text(f"✅ Completed group {total_groups}/{total_expected_groups}: {group_name} | Waiting 5 seconds...")
        
        # Sleep for 5 seconds to prevent timeout issues
        time.sleep(5)
    
    return csv_files, total_groups, total_files

def main():
    st.set_page_config(
        page_title="Excel Splitter Tool",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel File Splitter")
    st.markdown("Upload an Excel file and split it into multiple CSV files based on grouping criteria")
    
    # File upload
    st.header("1. Upload Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file", 
        type=['xlsx', 'xls'],
        help="Upload your Excel file (up to 200MB)"
    )
    
    if uploaded_file is not None:
        try:
            # Load the Excel file
            with st.spinner("Loading Excel file..."):
                df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File loaded successfully! Shape: {df.shape[0]:,} rows × {df.shape[1]} columns")
            
            # Display basic info about the file
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", f"{df.shape[0]:,}")
            with col2:
                st.metric("Total Columns", df.shape[1])
            with col3:
                st.metric("File Size", f"{uploaded_file.size / (1024*1024):.2f} MB")
            
            # Preview data
            st.header("2. Data Preview")
            with st.expander("👀 View first 10 rows", expanded=False):
                st.dataframe(df.head(10))
            
            # Column selection for grouping
            st.header("3. Select Grouping Columns")
            st.markdown("Choose one or more columns to group by. Each unique combination will create separate CSV files.")
            
            available_columns = df.columns.tolist()
            group_columns = st.multiselect(
                "Select columns for grouping:",
                available_columns,
                help="You can select multiple columns. Each unique combination of values will create a separate file."
            )
            
            if group_columns:
                # Show preview of unique combinations
                if len(group_columns) == 1:
                    unique_combos = df[group_columns[0]].nunique()
                    preview_combos = df[group_columns[0]].unique()[:10]
                else:
                    unique_combos = df[group_columns].drop_duplicates().shape[0]
                    preview_combos = df[group_columns].drop_duplicates().head(10)
                
                st.info(f"📈 This will create approximately **{unique_combos}** different groups")
                
                with st.expander(f"Preview of unique combinations (showing first 10)", expanded=False):
                    if len(group_columns) == 1:
                        st.write("Unique values:")
                        for combo in preview_combos:
                            st.write(f"• {combo}")
                    else:
                        st.dataframe(preview_combos)
            
            # Column selection for output
            st.header("4. Select Columns for Output")
            st.markdown("Choose which columns to include in the generated CSV files.")
            
            selected_columns = st.multiselect(
                "Select columns to include in CSV files:",
                available_columns,
                default=available_columns,
                help="Select the columns you want to include in the output CSV files"
            )
            
            # Additional settings
            st.header("5. Additional Settings")
            col1, col2 = st.columns(2)
            
            with col1:
                max_rows_per_file = st.number_input(
                    "Maximum rows per CSV file:",
                    min_value=50,
                    max_value=10000,
                    value=200,
                    step=50,
                    help="If a group has more rows than this limit, it will be split into multiple files"
                )
            
            with col2:
                include_timestamp = st.checkbox(
                    "Include timestamp in filenames",
                    value=False,
                    help="Add current timestamp to avoid filename conflicts"
                )
            
            # Process button
            if group_columns and selected_columns:
                st.header("6. Generate CSV Files")
                
                if st.button("🚀 Generate CSV Files", type="primary"):
                    # Create progress indicators
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Estimate processing time
                    if len(group_columns) == 1:
                        estimated_groups = df[group_columns[0]].nunique()
                    else:
                        estimated_groups = df[group_columns].drop_duplicates().shape[0]
                    
                    estimated_time_minutes = (estimated_groups * 5) / 60
                    st.info(f"⏱️ Estimated processing time: {estimated_time_minutes:.1f} minutes ({estimated_groups} groups × 5 seconds each)")
                    
                    with st.spinner("Processing data and creating CSV files..."):
                        try:
                            csv_files, total_groups, total_files = process_excel_data(
                                df, group_columns, selected_columns, max_rows_per_file, 
                                progress_bar, status_text
                            )
                            
                            # Clear progress indicators
                            progress_bar.empty()
                            status_text.empty()
                            
                            if include_timestamp:
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                csv_files = [(f"{timestamp}_{filename}", content) for filename, content in csv_files]
                            
                            # Create ZIP file
                            with st.spinner("Creating ZIP file for download..."):
                                zip_buffer = create_zip_download(csv_files)
                            
                            # Success message
                            st.success(f"✅ Successfully created {total_files} CSV files from {total_groups} unique groups!")
                            
                            # Summary statistics
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Groups Created", total_groups)
                            with col2:
                                st.metric("CSV Files Generated", total_files)
                            with col3:
                                st.metric("ZIP File Size", f"{len(zip_buffer.getvalue()) / (1024*1024):.2f} MB")
                            
                            # Download button
                            st.download_button(
                                label="📥 Download All CSV Files (ZIP)",
                                data=zip_buffer.getvalue(),
                                file_name=f"split_csvs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                                mime="application/zip",
                                type="primary"
                            )
                            
                            # Show file list
                            with st.expander("📁 List of generated files", expanded=False):
                                for i, (filename, _) in enumerate(csv_files, 1):
                                    st.write(f"{i}. {filename}")
                                    
                        except Exception as e:
                            # Clear progress indicators on error
                            progress_bar.empty()
                            status_text.empty()
                            st.error(f"❌ Error processing data: {str(e)}")
                            st.error("Please check your data and try again.")
            else:
                if not group_columns:
                    st.warning("⚠️ Please select at least one grouping column.")
                if not selected_columns:
                    st.warning("⚠️ Please select at least one column for output.")
        
        except Exception as e:
            st.error(f"❌ Error loading file: {str(e)}")
            st.error("Please make sure you uploaded a valid Excel file.")
    
    else:
        st.info("👆 Please upload an Excel file to get started.")
    
    # Instructions
    with st.sidebar:
        st.header("📖 How to Use")
        st.markdown("""
        1. **Upload** your Excel file (xlsx or xls)
        2. **Select grouping columns** - each unique combination creates a separate CSV
        3. **Choose output columns** - select which columns to include in CSV files
        4. **Set max rows per file** - files larger than this will be split
        5. **Generate and download** the ZIP file containing all CSV files
        """)
        
        st.header("💡 Tips")
        st.markdown("""
        - **Multiple grouping columns**: Creates files for each unique combination
        - **Large groups**: Automatically split into smaller files
        - **File naming**: Based on group values (special characters removed)
        - **Local processing**: Everything runs on your computer
        """)
        
        st.header("⚙️ Technical Info")
        st.markdown(f"""
        - **Max file size**: ~200MB Excel files
        - **Processing**: 5-second delay between groups (for cloud stability)
        - **Memory efficient**: Processes data in chunks
        - **Output format**: CSV files in ZIP archive
        - **Cloud optimized**: Prevents timeout issues
        """)

        st.header("⚠️ Cloud Processing Notes")
        st.markdown("""
        - **Processing time**: ~5 seconds per unique group
        - **Large datasets**: May take several minutes
        - **Progress tracking**: Real-time status updates
        - **Timeout prevention**: Built-in delays for stability
        """)

if __name__ == "__main__":
    main()