import streamlit as st
import io
import time
from datetime import datetime
from processor import process_docx

# =========================
# GIAO DIỆN (UI/UX)
# =========================

st.set_page_config(
    page_title="Word Processor Pro",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "# Word Processor Pro\nProfessional document optimization."
    }
)

# Custom CSS for Premium Look
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }

    /* Hide Streamlit branding and GitHub link */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Center the main container */
    .main .block-container {
        max-width: 800px;
        padding-top: 5rem;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        margin-top: 2rem;
        margin-bottom: 2rem;
    }

    /* Title Styling */
    h1 {
        color: #1e3a8a;
        font-family: 'Outfit', sans-serif;
        font-weight: 800;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        color: #64748b;
        text-align: center;
        font-size: 1.1rem;
        margin-bottom: 3rem;
    }

    /* File Uploader Container */
    .stFileUploader {
        padding: 2rem;
        border: 2px dashed #3b82f6;
        border-radius: 15px;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: #1d4ed8;
        background: rgba(59, 130, 246, 0.05);
    }

    /* Button Styling */
    .stButton>button {
        width: 100%;
        background-color: #3b82f6;
        color: white;
        border-radius: 10px;
        padding: 0.75rem;
        font-weight: 600;
        border: none;
        transition: transform 0.2s ease;
    }
    
    .stButton>button:hover {
        background-color: #2563eb;
        transform: translateY(-2px);
    }

    /* Success Message */
    .stSuccess {
        background-color: #dcfce7;
        color: #166534;
        border: 1px solid #bbf7d0;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

def word_processor_tab():
    # Content of Word Processor Pro
    st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>Word Processor Pro</h2>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Professional Word document optimization and standardization</p>", unsafe_allow_html=True)

    # Settings
    auto_process = st.checkbox("Auto-process after upload", value=True, key="auto_process_check")

    # File Uploader
    uploaded_file = st.file_uploader(
        "Drag and drop Word file (.docx) here",
        type=["docx"],
        help="Only .docx format is supported",
        key="word_proc_uploader"
    )

    if uploaded_file is not None:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        # Info about selected file
        st.info(f"Selected: **{uploaded_file.name}**")

        # Start button (Manual trigger)
        start_manual = st.button("🚀 Start Processing", key="btn_process")
        
        # Logic to decide if we should run processing
        # Run if: manual click OR (auto-process is ON and this file hasn't been processed yet)
        should_run = start_manual or (auto_process and st.session_state.get('last_processed_id') != file_id)

        if should_run:
            # Update state to prevent double execution
            st.session_state['last_processed_id'] = file_id
            
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # UX Steps
                status_text.text("Loading document...")
                time.sleep(0.5)
                progress_bar.progress(20)

                status_text.text("Normalizing headings and fonts...")
                # Important: Read file into memory
                input_data = uploaded_file.getvalue()
                input_stream = io.BytesIO(input_data)
                progress_bar.progress(50)

                # Core processing
                processed_stream = process_docx(input_stream)
                
                status_text.text("Finalizing file...")
                progress_bar.progress(90)
                time.sleep(0.3)
                progress_bar.progress(100)
                
                # Store results in session state for persistence across reruns
                st.session_state['processed_data'] = processed_stream.getvalue()
                time_tag = datetime.now().strftime("%H%M")
                base_name = uploaded_file.name.rsplit('.', 1)[0]
                st.session_state['output_filename'] = f"{base_name} edited {time_tag}.docx"
                
                st.success("✅ Processing complete!")

            except Exception as e:
                st.error(f"❌ An error occurred during processing: {str(e)}")
                # Reset state on error so user can retry
                st.session_state['last_processed_id'] = None

        # Display download button if we have processed data for the current file
        if st.session_state.get('last_processed_id') == file_id and 'processed_data' in st.session_state:
            st.download_button(
                label="📥 Download Processed File",
                data=st.session_state['processed_data'],
                file_name=st.session_state['output_filename'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="btn_download"
            )

def main():
    # Application Title
    st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>Ultimate Tool Hub</h1>", unsafe_allow_html=True)
    
    # Define Tabs
    tab_titles = ["📄 Word Processor Pro", "🛠️ Other Tools"]
    tabs = st.tabs(tab_titles)

    with tabs[0]:
        word_processor_tab()

    with tabs[1]:
        st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>Coming Soon</h2>", unsafe_allow_html=True)
        st.info("We are developing new tools to make your work more efficient. Stay tuned!")

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #94a3b8; font-size: 0.8rem;'>"
        "Efficiency Optimization System."
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
