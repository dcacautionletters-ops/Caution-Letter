import streamlit as st
import pandas as pd
from fpdf import FPDF

def generate_pdf(df):
    # Initialize A4 PDF
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # EXACT DIMENSIONS (mm)
    label_w = 100  # 10 cm
    label_h = 43   # 4.3 cm
    
    # LAYOUT SPACING (mm)
    top_margin = 10    # 1 cm from top
    right_margin = 7   # 0.7 cm from right
    # (210mm A4 width - 200mm labels width - 7mm right margin) = 3mm left
    left_margin = 3    
    
    for i, (_, row) in enumerate(df.iterrows()):
        col = i % 2
        line = (i // 2) % 6  # 6 rows per page
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        x = left_margin + (col * label_w)
        y = top_margin + (line * label_h)
        
        # Draw Light Grey Border
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        # Position cursor inside the label
        pdf.set_xy(x + 5, y + 5)
        
        # Function to strip characters that crash PDF engines
        def clean(val):
            if pd.isna(val): return ""
            return str(val).encode('ascii', 'ignore').decode('ascii')

        # Data Mapping based on your files
        father = clean(row.get('FATHER', ''))
        student = clean(row.get('NAME', ''))
        address = clean(row.get('ADDRESS', ''))
        roll = clean(row.get('ROLL NUMBER', ''))
        ph1 = clean(row.get('STUDENT PHONE', ''))
        ph2 = clean(row.get('PARENT PHONE', ''))
        contact = f"{ph1}/{ph2}".strip("/")
        
        # --- PRINT CONTENT ---
        # Header
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Address Details
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 5)
        
        label_body = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {student}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        pdf.multi_cell(label_w - 10, 5, label_body)
        
    # Handle both fpdf and fpdf2 output types to prevent encoding errors
    pdf_out = pdf.output(dest='S')
    if isinstance(pdf_out, str):
        return pdf_out.encode('latin-1')
    return pdf_out

# --- STREAMLIT INTERFACE ---
st.title("🏷️ 2x6 Professional Label Generator")

file_c = st.file_uploader("Upload Caution Letter File (Col B must be Roll No)", type=['xlsx', 'csv'])
file_m = st.file_uploader("Upload Master Database (Must have 'ROLL NUMBER' column)", type=['xlsx', 'csv'])

if file_c and file_m:
    try:
        # Load data
        df_c = pd.read_csv(file_c) if file_c.name.endswith('csv') else pd.read_excel(file_c)
        df_m = pd.read_csv(file_m) if file_m.name.endswith('csv') else pd.read_excel(file_m)

        # Matching Logic
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.strip()
        
        df_final = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]

        if st.button("Generate 12 Labels Per Sheet"):
            if not df_final.empty:
                pdf_bytes = generate_pdf(df_final)
                st.download_button(
                    label="📥 Download PDF",
                    data=pdf_bytes,
                    file_name="Student_Address_Labels.pdf",
                    mime="application/pdf"
                )
                st.success(f"Labels generated for {len(df_final)} students.")
            else:
                st.error("No matches found. Ensure 'ROLL NUMBER' column exists in Master file.")
                
    except Exception as e:
        st.error(f"Error: {e}")
