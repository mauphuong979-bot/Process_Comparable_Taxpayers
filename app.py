import streamlit as st
from word_tab import word_processor_tab
from matrix_tab import render_matrix_processor_tab

# =========================
# GIAO DIỆN (UI/UX)
# =========================

st.set_page_config(
    page_title="Ultimate Tool Hub",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
    menu_items={
        'Get Help': None,
        'Report a bug': None,
        'About': "# Ultimate Tool Hub\nProfessional document optimization and data processing system."
    }
)

# Custom CSS for Premium/Modern Look
st.markdown("""
    <style>
    /* Google Fonts Import */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&family=Outfit:wght@400;700;800&display=swap');

    /* Global Background */
    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
    }

    /* Modern Scrollbar */
    ::-webkit-scrollbar {
        width: 8px;
        height: 8px;
    }
    ::-webkit-scrollbar-track {
        background: rgba(0, 0, 0, 0.05);
    }
    ::-webkit-scrollbar-thumb {
        background: rgba(59, 130, 246, 0.3);
        border-radius: 10px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: rgba(59, 130, 246, 0.5);
    }

    /* Hide Streamlit default components */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Main Container Styling - Premium Glassmorphism */
    [data-testid="stMainBlockContainer"] {
        max-width: 1200px;
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        margin: auto;
        background: rgba(255, 255, 255, 0.75);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border: 1px solid rgba(255, 255, 255, 0.4);
        border-radius: 24px;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.08);
        margin-top: 1.5rem;
        margin-bottom: 1.5rem;
    }

    /* Title & Typography Styling */
    h1 {
        color: #1e3a8a;
        font-family: 'Outfit', sans-serif;
        font-weight: 800;
        text-align: center;
        letter-spacing: -0.5px;
        margin-bottom: 0.5rem;
    }
    
    .subtitle {
        color: #475569;
        text-align: center;
        font-family: 'Inter', sans-serif;
        font-size: 1.1rem;
        margin-bottom: 2.5rem;
        opacity: 0.8;
    }

    /* Tab Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        background-color: transparent;
    }

    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: rgba(255, 255, 255, 0.5);
        border-radius: 12px;
        color: #64748b;
        font-weight: 600;
        padding: 10px 25px;
        border: 1px solid transparent;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }

    .stTabs [data-baseweb="tab"]:hover {
        background-color: rgba(59, 130, 246, 0.1);
        color: #3b82f6;
    }

    .stTabs [aria-selected="true"] {
        background-color: #3b82f6 !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    }

    /* File Uploader Premium Styling */
    [data-testid="stFileUploader"] {
        padding: 1rem;
        border: 2px dashed rgba(59, 130, 246, 0.3);
        border-radius: 18px;
        background: rgba(255, 255, 255, 0.4);
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: #3b82f6;
        background: rgba(59, 130, 246, 0.05);
    }

    /* Button Styling - Modern Hover */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        border-radius: 12px;
        padding: 0.7rem 1.5rem;
        font-weight: 600;
        border: none;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .stButton>button:hover {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        box-shadow: 0 8px 15px rgba(37, 99, 235, 0.3);
        transform: translateY(-2px);
    }

    .stButton>button:active {
        transform: translateY(0);
    }

    /* Message Styling */
    .stSuccess {
        border-radius: 12px;
        border: 1px solid #bbf7d0;
    }

    /* Responsive adjustments */
    @media (max-width: 768px) {
        [data-testid="stMainBlockContainer"] {
            padding: 1rem !important;
            margin-top: 0;
            border-radius: 0;
            max-width: 100%;
        }
    }
    </style>
""", unsafe_allow_html=True)

def main():
    """
    Main Application Entry Point
    """
    # 1. Application Header
    st.markdown("<h1>Ultimate Tool Hub</h1>", unsafe_allow_html=True)
    
    # 2. Tabs Navigation
    tab_titles = ["📄 Comparable Taxpayers Processor", "📊 Matrix Processor", "🛠️ Other Tools"]
    tabs = st.tabs(tab_titles)

    # 3. Tab Contents
    with tabs[0]:
        # Comparable Taxpayers Processor Tab
        word_processor_tab()

    with tabs[1]:
        # Matrix Processor Tab
        render_matrix_processor_tab()

    with tabs[2]:
        # Placeholder for future tools
        st.markdown("<h2 style='text-align: center; color: #1e3a8a;'>Coming Soon</h2>", unsafe_allow_html=True)
        st.info("We are developing new tools to make your work more efficient. Stay tuned!")

    # 4. Global Footer
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.markdown(
        "<div style='text-align: center; color: #64748b; font-size: 0.85rem; padding-bottom: 1rem;'>"
        "Professional Efficiency Optimization System &copy; 2024"
        "</div>", 
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
