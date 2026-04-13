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
)

# Custom CSS for Premium Look
st.markdown("""
    <style>
    /* Main Background */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    
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

def main():
    # Header Section
    st.markdown("<h1>Word Processor Pro</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Tối ưu hóa và chuẩn hóa tài liệu Word chuyên nghiệp</p>", unsafe_allow_html=True)

    # File Uploader
    uploaded_file = st.file_uploader(
        "Kéo và thả file Word (.docx) vào đây",
        type=["docx"],
        help="Chỉ hỗ trợ định dạng .docx"
    )

    if uploaded_file is not None:
        file_name = uploaded_file.name
        st.info(f"Đã chọn: **{file_name}**")

        if st.button("🚀 Bắt đầu xử lý"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                # Simulate steps for UX
                status_text.text("Đang tải tài liệu...")
                time.sleep(0.5)
                progress_bar.progress(20)

                status_text.text("Đang chuẩn hóa heading và font...")
                input_stream = io.BytesIO(uploaded_file.read())
                progress_bar.progress(50)

                # Core processing
                processed_stream = process_docx(input_stream)
                
                status_text.text("Đang hoàn tất lưu file...")
                progress_bar.progress(90)
                time.sleep(0.3)
                progress_bar.progress(100)
                
                # Success state
                st.success("✅ Xử lý hoàn tất!")
                
                # Preparation of download link
                time_tag = datetime.now().strftime("%H%M")
                base_name = file_name.rsplit('.', 1)[0]
                output_filename = f"{base_name} edited {time_tag}.docx"

                st.download_button(
                    label="📥 Tải file đã xử lý",
                    data=processed_stream,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra trong quá trình xử lý: {str(e)}")

    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #94a3b8; font-size: 0.8rem;'>"
        "Sử dụng công nghệ python-docx cho hiệu suất tối đa."
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
