import streamlit as st
import io
import time
from datetime import datetime
from processor import process_docx

def word_processor_tab():
    # Content of Comparable Taxpayers Processor
    st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>Comparable Taxpayers Processor</h2>", unsafe_allow_html=True)
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
