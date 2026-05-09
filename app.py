import streamlit as st
from word_tab import word_processor_tab
from matrix_tab import render_matrix_processor_tab

# =========================
# GIAO DIỆN (UI/UX)
# =========================

st.set_page_config(
    page_title="Comparable Taxpayers Processor",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "# Comparable Taxpayers Processor\nProfessional document optimization."
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

def main():
    # Application Title
    st.markdown("<h1 style='text-align: center; color: #1e3a8a;'>Ultimate Tool Hub</h1>", unsafe_allow_html=True)
    
    # Define Tabs
    tab_titles = ["📄 Comparable Taxpayers Processor", "📊 Matrix Processor", "🛠️ Other Tools"]
    tabs = st.tabs(tab_titles)

    with tabs[0]:
        word_processor_tab()

    with tabs[1]:
        render_matrix_processor_tab()

    with tabs[2]:
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
