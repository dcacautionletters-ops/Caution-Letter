import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

def generate_pdf(df):
    # A4 dimensions in mm: 210 x 297
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Label Specs: 10cm x 4.3cm
    label_w = 100
    label_h = 43
    
    # Margins from your requirement
    top_margin = 10    # 1 cm
    right_margin = 7   # 0.7 cm
    # We calculate left margin to balance the 0.7cm right margin on A4 (210mm)
    # (210 - (100*2) - 7) = 3mm left margin
    left_margin = 3 
    
    for i, row in df.iterrows():
        # 12 labels per page (2x6)
        col = i % 2
        line = (i // 2) % 6
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        x = left_margin + (col * label_w)
        y = top_margin + (line * label_h)
        
        # Draw Label Border
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        # Text Offsets
        pdf.set_xy(x + 5, y + 5)
        
        # From Header (Bold)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Body
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 5)
        
        # Join phone numbers logic
        ph_parent = str(row['PARENT PHONE']) if pd.notna(row['PARENT PHONE']) else ""
        ph_student = str(row['STUDENT PHONE']) if pd.notna(row['STUDENT PHONE']) else ""
        contact = f"{ph_student}/{ph_parent}".strip("/")
        
        label_text = (
            f"To, Shri/Smt. {row['FATHER']}\n"
            f"c/o: {row['NAME']}\n"
            f"Address: {row['ADDRESS']}\n"
            f"Contact: {contact}   ID: {row['ROLL NUMBER']}"
        )
        pdf.multi_cell(label_w - 10, 5, label_text)
        
    return pdf.output(dest='S').encode('latin-1')

st.title("Precise Label Generator (2x6)")

file_caution = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'])
file_master = st.file_uploader("Upload Student Master Data", type=['xlsx', 'csv'])

if file_caution and file_master:
    # Load Data
    df_c = pd.read_csv(file_caution) if file_caution.name.endswith('csv') else pd.read_excel(file_caution)
    df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)
    
    # Standardize Column Names for matching
    # Caution Letter: Roll No is Col B (Index 1), Name is Col C (Index 2)
    # Master Data: ROLL NUMBER is Col B, NAME is Col F
    
    # Filter Master data based on Caution Letter Roll Numbers
    caution_rolls = df_c.iloc[:, 1].dropna().astype(str).unique()
    
    # Exact Match Logic
    df_final = df_m[df_m['ROLL NUMBER'].astype(str).isin(caution_rolls)]
    
    if st.button("Generate 12-Label PDF"):
        if not df_final.empty:
            pdf_bytes = generate_pdf(df_final)
            st.download_button("Download PDF", pdf_bytes, "Labels.pdf", "application/pdf")
            st.success(f"Matched {len(df_final)} students from Caution List.")
        else:
            st.error("No matching Roll Numbers found between files.")
