import io
import streamlit as st
from thesis_formatter.formatter import format_document
import thesis_formatter.config as config

# 1. Page Configuration for premium design
st.set_page_config(
    page_title="Academic Thesis Formatter",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Custom Styling (vibrant titles, margins, sleek cards)
st.markdown("""
    <style>
    .main {
        background-color: #fcfcfc;
    }
    .title-text {
        text-align: center;
        color: #1e3d59;
        font-family: 'Inter', 'Outfit', sans-serif;
        font-weight: 800;
        margin-top: 10px;
        margin-bottom: 5px;
    }
    .subtitle-text {
        text-align: center;
        color: #17b978;
        font-family: 'Outfit', sans-serif;
        font-size: 1.1rem;
        margin-bottom: 30px;
    }
    .card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='title-text'>🎓 Thesis Formatter Tool</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle-text'>Make your academic document clean and format compliant instantly</p>", unsafe_allow_html=True)

# 2. Concurrency defensive check (Limit file size)
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# Sidebar Configuration Settings
st.sidebar.markdown("### Formatting Options")
st.sidebar.write("Customize formatting rules to apply to the body paragraphs:")
remove_shading = st.sidebar.checkbox("Remove Copy-Paste Shading & Highlights", value=True)
remove_bold = st.sidebar.checkbox("Remove Leftover Bold Words", value=True)
force_black = st.sidebar.checkbox("Force Black Headings", value=True)

st.sidebar.markdown("---")
st.sidebar.markdown("### Document Specifications")
st.sidebar.markdown(f"**Font family**: {config.BODY_FONT}")
st.sidebar.markdown(f"**Font size**: Arial 12pt")
st.sidebar.markdown(f"**Line spacing**: 1.5 Lines")
st.sidebar.markdown(f"**Alignment**: Justified")

# 3. File Uploader Box (Browse files / Drag-and-drop)
st.markdown("<div class='card'>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Choose a Word Document (.docx) to standardise:",
    type=["docx"],
    key="thesis_uploader"
)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded_file is not None:
    # Validate file size defensively
    if uploaded_file.size > MAX_FILE_SIZE:
        st.error("❌ The uploaded file is too large! Please upload a file smaller than 10MB to ensure server stability.")
    else:
        st.info("✓ Document uploaded successfully.")
        
        # Auto-run processing on upload/config change
        with st.spinner("Processing formatting... (Runs cleaning utilities, alignment, margins)"):
            try:
                # Apply User Config Checklist dynamically
                config.REMOVE_BODY_SHADING = remove_shading
                config.REMOVE_BODY_BOLD = remove_bold
                config.FORCE_HEADINGS_BLACK = force_black
                
                # Read uploaded file stream into memory
                input_stream = io.BytesIO(uploaded_file.read())
                output_stream = io.BytesIO()
                
                # Run formatting orchestrator in RAM
                report = format_document(
                    input_path=input_stream,
                    output_path=output_stream
                )
                
                # Move to start of stream for download trigger
                output_stream.seek(0)
                
                st.success("🎉 Formatting complete! (Please check the formatting once from your side after downloading)")
                
                # Display brief statistics card
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Paragraphs processed", report.para_count)
                with col2:
                    st.metric("Tables formatted", report.tables_formatted)
                with col3:
                    st.metric("Excess spaces removed", report.blanks_removed)
                
                st.markdown("---")
                
                # 4. Download Trigger (RAM stream download, automatically deleted on page refresh)
                st.download_button(
                    label="📥 Download Formatted DOCX File",
                    data=output_stream,
                    file_name=f"Formatted_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Error styling the document: {str(e)}")
                st.write("Please ensure the document is not corrupted or password-protected.")
