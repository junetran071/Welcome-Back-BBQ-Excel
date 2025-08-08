import streamlit as st
import pandas as pd
import io
from typing import Optional

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
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        # Clean and standardize column names
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
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
        
        # Column name configuration
        st.subheader("Column Settings")
        comparison_column = st.text_input(
            "Comparison Column Name",
            value="bronco id",
            help="The column name used to match records between files (case-insensitive)"
        )
        
        # Sheet name configuration
        st.subheader("Excel Sheet Settings")
        hrt_sheet = st.text_input("HRT Majors Sheet Name", value="Sheet1")
        bbq_sheet = st.text_input("BBQ Attendees Sheet Name", value="Sheet1")
        
        st.markdown("""
        <div class="info-box">
            <h5>💡 Tips:</h5>
            <ul>
                <li>Column names are automatically cleaned (spaces trimmed, converted to lowercase)</li>
                <li>Make sure both files have the same column name for comparison</li>
                <li>Default sheet name is "Sheet1"</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    # Main content area
    col1, col2 = st.columns(2)
    
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

    # Processing section
    if 'hrt_df' in locals() and 'bbq_df' in locals() and hrt_df is not None and bbq_df is not None:
        st.markdown('<h2 class="section-header">🔄 Process Data</h2>', unsafe_allow_html=True)
        
        # Validate both files have the required column
        hrt_valid = validate_dataframe(hrt_df, "HRT Majors", comparison_column)
        bbq_valid = validate_dataframe(bbq_df, "BBQ Attendees", comparison_column)
        
        if hrt_valid and bbq_valid:
            if st.button("🚀 Filter Non-HRT Attendees", type="primary", use_container_width=True):
                with st.spinner("Processing data..."):
                    try:
                        # Filter out non-HRT majors
                        non_hrt_attendees = bbq_df[~bbq_df[comparison_column].isin(hrt_df[comparison_column])]
                        
                        # Display results
                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.write("### 📊 Results:")
                        st.write(f"- **Total BBQ Attendees:** {len(bbq_df)}")
                        st.write(f"- **HRT Majors:** {len(hrt_df)}")
                        st.write(f"- **HRT Majors at BBQ:** {len(bbq_df) - len(non_hrt_attendees)}")
                        st.write(f"- **Non-HRT Attendees:** {len(non_hrt_attendees)}")
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
        else:
            st.markdown("""
            <div class="warning-box">
                <h4>⚠️ Cannot Process</h4>
                <p>Please ensure both files contain the specified comparison column before processing.</p>
            </div>
            """, unsafe_allow_html=True)

    # Instructions section
    with st.expander("📖 Detailed Instructions", expanded=False):
        st.markdown("""
        ### How to Use This App:
        
        1. **Prepare Your Files:**
           - Ensure both Excel files have a common column for comparison (default: "Bronco ID")
           - Files should be in .xlsx or .xls format
        
        2. **Configure Settings (Optional):**
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
        
        ### File Format Requirements:
        - **Excel files (.xlsx or .xls)**
        - **Both files must have a matching column** (e.g., "Bronco ID", "Student ID", etc.)
        - **Column names are case-insensitive** and automatically cleaned
        
        ### Troubleshooting:
        - If you get a "column not found" error, check the column names in your files
        - Make sure the comparison column exists in both files
        - Verify your sheet names are correct
        """)

if __name__ == "__main__":
    main()
