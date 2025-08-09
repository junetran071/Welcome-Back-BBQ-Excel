
import streamlit as st
import pandas as pd
import io
from typing import Optional

# Check for required dependencies
try:
    import openpyxl
except ImportError:
    st.error("""
    ❌ **Missing Dependency Error**
    
    The `openpyxl` library is required to read Excel files but is not installed.
    
    **To fix this issue:**
    1. Stop the Streamlit app (Ctrl+C)
    2. Run: `pip install openpyxl`
    3. Restart the app: `streamlit run hrt_filter_app.py`
    
    **Alternative:** Install all dependencies at once:
    `pip install streamlit pandas openpyxl xlrd`
    """)
    st.stop()

# Page configuration
st.set_page_config(
    page_title="HRT Major Filter App",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #2c3e50;
        margin: 1.5rem 0 1rem 0;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
    }
    .info-box {
        background-color: #e8f4fd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #3498db;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ffc107;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def load_excel_file(uploaded_file, sheet_name: str = "Sheet1") -> Optional[pd.DataFrame]:
    """Load Excel file and return DataFrame with cleaned column names."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine='openpyxl')
        # Clean and standardize column names
        df.columns = df.columns.str.strip().str.lower()
        return df
    except ImportError as e:
        st.error("""
        ❌ **Missing Dependency Error**
        
        Please install openpyxl: `pip install openpyxl`
        """)
        return None
    except FileNotFoundError:
        st.error(f"❌ Sheet '{sheet_name}' not found in the Excel file. Please check the sheet name.")
        return None
    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        st.info("💡 **Troubleshooting tips:**")
        st.write("- Make sure the file is a valid Excel file (.xlsx or .xls)")
        st.write("- Check that the sheet name is correct")
        st.write("- Ensure the file is not corrupted or password-protected")
        return None

def validate_dataframe(df: pd.DataFrame, file_name: str, required_column: str) -> bool:
    """Validate that the DataFrame contains the required column."""
    if df is None:
        return False
    
    if required_column not in df.columns:
        st.error(f"❌ '{required_column}' column not found in {file_name}")
        st.write("Available columns:", list(df.columns))
        return False
    
    return True

def create_download_link(df: pd.DataFrame, filename: str) -> bytes:
    """Create downloadable Excel file."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Non_HRT_Attendees')
    return output.getvalue()



# Main App
def main():
    # Header
    st.markdown('<h1 class="main-header">🎓 HRT Major Filter App</h1>', unsafe_allow_html=True)
    
    st.markdown("""
    <div class="info-box">
        <h4>📋 What does this app do?</h4>
        <p>This app helps you identify BBQ attendees who are <strong>NOT</strong> HRT (Hospitality, Recreation & Tourism) majors by comparing two Excel files:</p>
        <ul>
            <li><strong>HRT Majors file:</strong> Contains list of students who are HRT majors</li>
            <li><strong>BBQ Attendees file:</strong> Contains list of all BBQ attendees</li>
        </ul>
        <p>The app will filter out HRT majors from the BBQ attendees list and provide you with a downloadable Excel file of non-HRT attendees.</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar for configuration
    with st.sidebar:
        st.markdown('<h2 class="section-header">⚙️ Configuration</h2>', unsafe_allow_html=True)
        
        # Sheet name configuration
        st.subheader("Excel Sheet Settings")
        hrt_sheet = st.text_input("HRT Majors Sheet Name", value="Sheet1")
        bbq_sheet = st.text_input("BBQ Attendees Sheet Name", value="Sheet1")
        
        st.markdown("""
        <div class="info-box">
            <h5>💡 Tips:</h5>
            <ul>
                <li>Upload your files first to see available columns</li>
                <li>Column names are automatically cleaned (spaces trimmed, converted to lowercase)</li>
                <li>Default sheet name is "Sheet1"</li>
                <li>Select the columns you want to compare after uploading</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    # Main content area
    col1, col2 = st.columns(2)
    
    # Initialize dataframes
    hrt_df = None
    bbq_df = None
    
    # File upload mode
    with col1:
        st.markdown('<h2 class="section-header">📁 Upload HRT Majors File</h2>', unsafe_allow_html=True)
        hrt_file = st.file_uploader(
            "Choose HRT Majors Excel file",
            type=['xlsx', 'xls'],
            key="hrt_file",
            help="Upload the Excel file containing HRT majors data"
        )
        
        if hrt_file is not None:
            with st.spinner("Loading HRT Majors file..."):
                hrt_df = load_excel_file(hrt_file, hrt_sheet)
            
            if hrt_df is not None:
                st.success(f"✅ Loaded {len(hrt_df)} HRT majors")
                
                # Show preview
                with st.expander("Preview HRT Majors Data"):
                    st.dataframe(hrt_df.head(), use_container_width=True)
                    st.write(f"**Columns:** {list(hrt_df.columns)}")

    with col2:
        st.markdown('<h2 class="section-header">📁 Upload BBQ Attendees File</h2>', unsafe_allow_html=True)
        bbq_file = st.file_uploader(
            "Choose BBQ Attendees Excel file",
            type=['xlsx', 'xls'],
            key="bbq_file",
            help="Upload the Excel file containing BBQ attendees data"
        )
        
        if bbq_file is not None:
            with st.spinner("Loading BBQ Attendees file..."):
                bbq_df = load_excel_file(bbq_file, bbq_sheet)
            
            if bbq_df is not None:
                st.success(f"✅ Loaded {len(bbq_df)} BBQ attendees")
                
                # Show preview
                with st.expander("Preview BBQ Attendees Data"):
                    st.dataframe(bbq_df.head(), use_container_width=True)
                    st.write(f"**Columns:** {list(bbq_df.columns)}")

    # Column Selection Section (only show when data is loaded)
    if hrt_df is not None and bbq_df is not None:
        st.markdown('<h2 class="section-header">🎯 Column Selection for Comparison</h2>', unsafe_allow_html=True)
        
        col_sel1, col_sel2 = st.columns(2)
        
        with col_sel1:
            st.subheader("🎓 HRT Majors File")
            hrt_column_options = list(hrt_df.columns)
            
            # Try to find a default column (bronco id, id, student id, etc.)
            default_hrt_idx = 0
            for i, col in enumerate(hrt_column_options):
                if any(keyword in col.lower() for keyword in ['bronco id', 'id', 'student id', 'student_id']):
                    default_hrt_idx = i
                    break
            
            hrt_comparison_column = st.selectbox(
                "Select comparison column from HRT Majors:",
                options=hrt_column_options,
                index=default_hrt_idx,
                key="hrt_column_select",
                help="Choose the column that contains unique identifiers (e.g., Student ID, Bronco ID)"
            )
            
            # Show sample values
            if hrt_comparison_column:
                sample_values = hrt_df[hrt_comparison_column].dropna().head(3).tolist()
                st.write(f"**Sample values:** {sample_values}")
        
        with col_sel2:
            st.subheader("🍖 BBQ Attendees File")
            bbq_column_options = list(bbq_df.columns)
            
            # Try to find a default column that matches or is similar to HRT column
            default_bbq_idx = 0
            if hrt_comparison_column:
                for i, col in enumerate(bbq_column_options):
                    if (col.lower() == hrt_comparison_column.lower() or 
                        any(keyword in col.lower() for keyword in ['bronco id', 'id', 'student id', 'student_id'])):
                        default_bbq_idx = i
                        break
            
            bbq_comparison_column = st.selectbox(
                "Select comparison column from BBQ Attendees:",
                options=bbq_column_options,
                index=default_bbq_idx,
                key="bbq_column_select",
                help="Choose the column that contains the same type of identifiers as in HRT Majors file"
            )
            
            # Show sample values
            if bbq_comparison_column:
                sample_values = bbq_df[bbq_comparison_column].dropna().head(3).tolist()
                st.write(f"**Sample values:** {sample_values}")
        
        # Show comparison summary
        if hrt_comparison_column and bbq_comparison_column:
            st.markdown("""
            <div class="info-box">
                <h4>🔍 Comparison Setup:</h4>
                <ul>
                    <li><strong>HRT Majors column:</strong> {}</li>
                    <li><strong>BBQ Attendees column:</strong> {}</li>
                    <li><strong>Comparison method:</strong> Find BBQ attendees whose {} is NOT in the HRT Majors {} list</li>
                </ul>
            </div>
            """.format(hrt_comparison_column, bbq_comparison_column, bbq_comparison_column, hrt_comparison_column), 
            unsafe_allow_html=True)
            
            # Data type validation
            hrt_sample = hrt_df[hrt_comparison_column].dropna().iloc[0] if not hrt_df[hrt_comparison_column].dropna().empty else None
            bbq_sample = bbq_df[bbq_comparison_column].dropna().iloc[0] if not bbq_df[bbq_comparison_column].dropna().empty else None
            
            if hrt_sample is not None and bbq_sample is not None:
                hrt_type = type(hrt_sample).__name__
                bbq_type = type(bbq_sample).__name__
                
                if hrt_type != bbq_type:
                    st.warning(f"⚠️ **Data Type Mismatch:** HRT column contains {hrt_type} while BBQ column contains {bbq_type}. This might affect comparison accuracy.")
                else:
                    st.success(f"✅ **Data Types Match:** Both columns contain {hrt_type} values.")
    
    else:
        # Show placeholder when no data is loaded
        st.markdown("""
        <div class="info-box">
            <h4>📋 Column Selection</h4>
            <p>Upload both Excel files to see available columns for comparison. The app will:</p>
            <ul>
                <li>Automatically detect potential ID columns</li>
                <li>Allow you to select the exact columns to compare</li>
                <li>Show sample values to verify your selection</li>
                <li>Validate data types for accurate comparison</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    # Processing section
    if hrt_df is not None and bbq_df is not None:
        st.markdown('<h2 class="section-header">🔄 Process Data</h2>', unsafe_allow_html=True)
        
        # Get selected columns from session state
        hrt_comparison_column = st.session_state.get("hrt_column_select")
        bbq_comparison_column = st.session_state.get("bbq_column_select")
        
        # Validate both files have the required columns
        if hrt_comparison_column and bbq_comparison_column:
            hrt_valid = hrt_comparison_column in hrt_df.columns
            bbq_valid = bbq_comparison_column in bbq_df.columns
            
            if not hrt_valid:
                st.error(f"❌ Column '{hrt_comparison_column}' not found in HRT Majors file")
            if not bbq_valid:
                st.error(f"❌ Column '{bbq_comparison_column}' not found in BBQ Attendees file")
            
            if hrt_valid and bbq_valid:
                if st.button("🚀 Filter Non-HRT Attendees", type="primary", use_container_width=True):
                    with st.spinner("Processing data..."):
                        try:
                            # Filter out non-HRT majors
                            non_hrt_attendees = bbq_df[~bbq_df[bbq_comparison_column].isin(hrt_df[hrt_comparison_column])]
                            
                            # Display results
                            st.markdown('<div class="success-box">', unsafe_allow_html=True)
                            st.write("### 📊 Results:")
                            st.write(f"- **Total BBQ Attendees:** {len(bbq_df)}")
                            st.write(f"- **Total HRT Majors:** {len(hrt_df)}")
                            st.write(f"- **HRT Majors at BBQ:** {len(bbq_df) - len(non_hrt_attendees)}")
                            st.write(f"- **Non-HRT Attendees:** {len(non_hrt_attendees)}")
                            st.write(f"- **Comparison Method:** Comparing '{bbq_comparison_column}' with '{hrt_comparison_column}'")
                            
                            # Show which HRT majors attended
                            hrt_at_bbq = bbq_df[bbq_df[bbq_comparison_column].isin(hrt_df[hrt_comparison_column])]
                            if len(hrt_at_bbq) > 0:
                                # Try to get name column for display
                                name_col = None
                                for col in bbq_df.columns:
                                    if 'name' in col.lower():
                                        name_col = col
                                        break
                                
                                if name_col:
                                    hrt_names = hrt_at_bbq[name_col].tolist()
                                    st.write(f"- **HRT Majors who attended BBQ:** {', '.join(map(str, hrt_names))}")
                                else:
                                    st.write(f"- **HRT Major IDs at BBQ:** {', '.join(map(str, hrt_at_bbq[bbq_comparison_column].tolist()))}")
                            
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                            # Show filtered data
                            if len(non_hrt_attendees) > 0:
                                st.subheader("👥 Non-HRT BBQ Attendees")
                                st.dataframe(non_hrt_attendees, use_container_width=True)
                                
                                # Download button
                                excel_data = create_download_link(non_hrt_attendees, "Non_HRT_BBQ_Attendees.xlsx")
                                st.download_button(
                                    label="📥 Download Non-HRT Attendees Excel File",
                                    data=excel_data,
                                    file_name="Non_HRT_BBQ_Attendees.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary",
                                    use_container_width=True
                                )
                            else:
                                st.info("🎉 All BBQ attendees are HRT majors!")
                                
                        except Exception as e:
                            st.error(f"❌ Error processing data: {str(e)}")
                            st.write("**Debug info:**")
                            st.write(f"- HRT column: {hrt_comparison_column}")
                            st.write(f"- BBQ column: {bbq_comparison_column}")
                            st.write(f"- HRT shape: {hrt_df.shape}")
                            st.write(f"- BBQ shape: {bbq_df.shape}")
            else:
                st.markdown("""
                <div class="warning-box">
                    <h4>⚠️ Cannot Process</h4>
                    <p>Please ensure both files contain the selected comparison columns before processing.</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="warning-box">
                <h4>⚠️ Column Selection Required</h4>
                <p>Please select comparison columns from both files in the "Column Selection" section above.</p>
            </div>
            """, unsafe_allow_html=True)

    # Instructions section
    with st.expander("📖 Detailed Instructions", expanded=False):
        st.markdown("""
        ### How to Use This App:
        
        #### 📁 Step-by-Step Instructions:
        1. **Prepare Your Files:**
           - Ensure both Excel files have columns with matching data for comparison
           - Files should be in .xlsx or .xls format
        
        2. **Upload Files:**
           - Upload the HRT Majors file on the left
           - Upload the BBQ Attendees file on the right
           - Preview the data to ensure it loaded correctly
        
        3. **Select Comparison Columns:**
           - Choose the appropriate column from each file for comparison
           - The app will auto-detect potential ID columns
           - Verify sample values match the expected format
           - Check that data types are compatible
        
        4. **Configure Settings (Optional):**
           - Modify sheet names if your data is not in "Sheet1"
        
        5. **Process Data:**
           - Click "Filter Non-HRT Attendees" to process the data
           - Review the results summary
           - Download the filtered Excel file
        
        ### Column Selection Features:
        - **Auto-detection:** App automatically suggests likely ID columns
        - **Sample values:** Preview data to verify your selection
        - **Data type validation:** Ensures compatible data types for accurate comparison
        - **Flexible matching:** Can use different column names (e.g., "Student ID" vs "Bronco ID")
        
        ### File Format Requirements:
        - **Excel files (.xlsx or .xls)**
        - **Both files must have columns with comparable data** (e.g., "Bronco ID", "Student ID", etc.)
        - **Column names are case-insensitive** and automatically cleaned
           - Use the sidebar to change the comparison column name if needed
           - Modify sheet names if your data is not in "Sheet1"
        
        3. **Upload Files:**
           - Upload the HRT Majors file on the left
           - Upload the BBQ Attendees file on the right
           - Preview the data to ensure it loaded correctly
        
        4. **Process Data:**
           - Click "Filter Non-HRT Attendees" to process the data
           - Review the results summary
           - Download the filtered Excel file
        
        ### Sample Data Information:
        The demo data includes:
        - **BBQ Attendees:** 10 students with Name, cpp.edu email, and Bronco ID
        - **HRT Majors:** 10 students with Last name, First name, Email, and Bronco ID
        - **Expected Result:** 5 non-HRT attendees (Newman, Taylor, Zara, Wendy, Leo)
        
        ### File Format Requirements:
        - **Excel files (.xlsx or .xls)**
        - **Both files must have a matching column** (e.g., "Bronco ID", "Student ID", etc.)
        - **Column names are case-insensitive** and automatically cleaned
        
        ### Troubleshooting:
        - If you get a "column not found" error, check the column names in your files
        - Make sure the comparison column exists in both files
        - Verify your sheet names are correct
        - Try Demo Mode first to understand how the app works
        """)

    # Footer with additional info
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🎓 <strong>HRT Major Filter App</strong> | Built with Streamlit | 
        Upload your Excel files to get started</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
