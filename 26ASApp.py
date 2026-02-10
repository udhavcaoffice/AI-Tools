import streamlit as st
import pandas as pd
import io
import re
from pdf2image import convert_from_bytes
import pytesseract
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fuzzywuzzy import fuzz

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="AI Tools | Udhav Agarwalla & Co",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ==========================================
# CUSTOM CSS - PROFESSIONAL BRANDING
# ==========================================
st.markdown("""
<style>
    /* Import Inter Font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
    
    /* Global Styles */
    * {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }
    
    /* Hide Streamlit Branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Main App Background */
    .stApp {
        background-color: #ffffff;
    }
    
    /* Professional Header */
    .brand-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
        padding: 2.5rem 3rem;
        margin: -5rem -5rem 3rem -5rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }
    
    .company-name {
        font-size: 2.5rem;
        font-weight: 900;
        color: white;
        text-transform: uppercase;
        letter-spacing: -0.02em;
        margin: 0;
        line-height: 1.1;
    }
    
    .company-tagline {
        font-size: 0.75rem;
        color: rgba(255,255,255,0.85);
        text-transform: uppercase;
        letter-spacing: 0.3em;
        font-weight: 700;
        margin-top: 0.5rem;
    }
    
    /* Hero Section */
    .hero-section {
        text-align: center;
        padding: 2.5rem 1rem;
        max-width: 900px;
        margin: 0 auto 3rem;
    }
    
    .hero-title {
        font-size: 2.75rem;
        font-weight: 900;
        color: #1e3a8a;
        margin-bottom: 1rem;
        line-height: 1.2;
    }
    
    .hero-description {
        font-size: 1.125rem;
        color: #64748b;
        line-height: 1.7;
        font-weight: 500;
    }
    
    /* Tool Categories */
    .category-title {
        font-size: 1.5rem;
        font-weight: 900;
        color: #1e3a8a;
        margin: 2rem 0 1rem 0;
        text-transform: uppercase;
        letter-spacing: -0.01em;
    }
    
    /* Buttons */
    .stButton > button {
        background: #1e3a8a !important;
        color: white !important;
        font-weight: 700 !important;
        padding: 0.875rem 2.5rem !important;
        border-radius: 10px !important;
        border: none !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
        font-size: 0.875rem !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 14px rgba(30, 58, 138, 0.25) !important;
    }
    
    .stButton > button:hover {
        background: #1e40af !important;
        box-shadow: 0 6px 20px rgba(30, 58, 138, 0.35) !important;
        transform: translateY(-2px) !important;
    }
    
    .stDownloadButton > button {
        background: #059669 !important;
        color: white !important;
        font-weight: 700 !important;
        padding: 0.875rem 2.5rem !important;
        border-radius: 10px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.05em !important;
        font-size: 0.875rem !important;
    }
    
    .stDownloadButton > button:hover {
        background: #047857 !important;
        transform: translateY(-2px) !important;
    }
    
    /* File Uploader */
    .stFileUploader {
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 2rem;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: #1e3a8a;
        background: #f1f5f9;
    }
    
    /* Info/Success/Error Messages */
    .stAlert {
        border-radius: 12px !important;
        border-left: 4px solid #1e3a8a !important;
        font-weight: 500 !important;
        padding: 1rem 1.5rem !important;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        background: transparent;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: white;
        border: 2px solid #e2e8f0;
        border-radius: 10px;
        padding: 1rem 2rem;
        font-weight: 700;
        color: #64748b;
        transition: all 0.3s ease;
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        border-color: #cbd5e1;
        background: #f8fafc;
    }
    
    .stTabs [aria-selected="true"] {
        background: #1e3a8a !important;
        color: white !important;
        border-color: #1e3a8a !important;
    }
    
    /* Headers */
    h1, h2, h3, h4 {
        color: #1e3a8a !important;
        font-weight: 900 !important;
    }
    
    /* DataFrames */
    .dataframe {
        border: 2px solid #e2e8f0 !important;
        border-radius: 10px !important;
    }
    
    /* Progress Bar */
    .stProgress > div > div {
        background-color: #1e3a8a !important;
    }
    
    /* Slider */
    .stSlider > div > div > div {
        background-color: #1e3a8a !important;
    }
    
    /* Footer */
    .custom-footer {
        text-align: center;
        padding: 3rem 0 2rem 0;
        color: #94a3b8;
        font-size: 0.875rem;
        margin-top: 5rem;
        border-top: 2px solid #e2e8f0;
    }
    
    .footer-link {
        color: #1e3a8a;
        text-decoration: none;
        font-weight: 600;
    }
    
    /* Coming Soon Badge */
    .coming-soon {
        display: inline-block;
        background: #fef3c7;
        color: #92400e;
        padding: 0.25rem 0.75rem;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-left: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# BRANDED HEADER
# ==========================================
st.markdown("""
<div class="brand-header">
    <div class="company-name">UDHAV AGARWALLA & CO</div>
    <div class="company-tagline">Chartered Accountants • AI-Powered Professional Tools</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# HERO SECTION
# ==========================================
st.markdown("""
<div class="hero-section">
    <h1 class="hero-title">Professional AI Tools for CAs</h1>
    <p class="hero-description">
        Streamline your audit, tax, and compliance processes with our advanced AI-powered utilities.
        <br>Built specifically for Chartered Accountants and financial professionals in India.
    </p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")

# ==========================================
# TOOL NAVIGATION TABS
# ==========================================
st.markdown('<p class="category-title">📋 Select Your Tool Category</p>', unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs([
    "📄 26AS Tools", 
    "📊 GST Tools (Coming Soon)", 
    "🔍 Tax Audit Tools (Coming Soon)"
])

# ==========================================
# TAB 1: 26AS TOOLS
# ==========================================
with tab1:
    st.markdown("### 🎯 26AS Processing Tools")
    st.markdown("Choose from our suite of 26AS tools:")
    
    # Sub-tabs for 26AS tools
    subtab1, subtab2, subtab3 = st.tabs([
        "1️⃣ PDF to Excel", 
        "2️⃣ Tally to Summary", 
        "3️⃣ Reconciliation"
    ])
    
    # ====================
    # TOOL 1: 26AS PDF TO EXCEL
    # ====================
    with subtab1:
        st.markdown("#### Convert 26AS PDF to Excel")
        st.info("📌 This tool uses OCR technology. Large PDF files may take 30-60 seconds to process.")
        
        uploaded_pdf = st.file_uploader(
            "Upload 26AS PDF File", 
            type="pdf", 
            key="pdf_upload",
            help="Select the 26AS PDF file downloaded from TRACES portal"
        )
        
        if uploaded_pdf:
            if st.button("🚀 Convert to Excel", key="convert_btn"):
                with st.spinner("🔄 Scanning PDF... Please wait (this may take 30-60 seconds)"):
                    try:
                        # Convert PDF bytes to images
                        images = convert_from_bytes(uploaded_pdf.read())
                        full_text = ""
                        
                        # Extract text from images
                        progress_bar = st.progress(0)
                        for i, image in enumerate(images):
                            text = pytesseract.image_to_string(image, config='--psm 4')
                            full_text += text + "\n"
                            progress_bar.progress((i + 1) / len(images))
                        
                        # Parse Text Logic
                        data = []
                        lines = full_text.split('\n')
                        tan_loose_pattern = re.compile(r'[A-Z]{4}[0-9OIl]{5}[A-Z]')
                        
                        for line in lines:
                            line = line.strip()
                            if len(line) < 15: continue
                            
                            match = tan_loose_pattern.search(line)
                            if match:
                                tan_code = match.group()
                                clean_line = re.sub(r'\s+', ' ', line)
                                parts = clean_line.split(' ')
                                
                                tan_idx = -1
                                for idx, part in enumerate(parts):
                                    if tan_code in part:
                                        tan_idx = idx
                                        break
                                
                                if tan_idx != -1:
                                    name_parts = []
                                    for j in range(tan_idx):
                                        w = parts[j]
                                        if len(w) > 1 and not re.match(r'^\d+$', w) and w.lower() not in ['sr', 'no']:
                                            name_parts.append(w)
                                    
                                    party_name = " ".join(name_parts)
                                    party_name = re.sub(r'^[^A-Z]+', '', party_name)
                                    
                                    amounts = []
                                    for token in parts[tan_idx+1:]:
                                        token_fix = token.replace('O','0').replace('o','0').replace('l','1').replace('I','1').replace('S','5')
                                        token_clean = re.sub(r'[^\d\.]', '', token_fix)
                                        if re.match(r'^\d+\.?\d{0,2}$', token_clean):
                                            try:
                                                val = float(token_clean)
                                                if val > 10:
                                                    amounts.append(val)
                                            except: pass
                                    
                                    final_tax = 0.0
                                    if len(amounts) >= 3:
                                        final_tax = amounts[1]
                                    elif len(amounts) == 2:
                                        final_tax = amounts[1]
                                    elif len(amounts) == 1:
                                        final_tax = amounts[0]
                                    
                                    if len(party_name) > 3 and final_tax > 0:
                                        data.append({
                                            'Name of Party': party_name,
                                            'Amount showing in 26AS': final_tax
                                        })
                        
                        df = pd.DataFrame(data)
                        
                        if not df.empty:
                            # Clean and Deduplicate
                
