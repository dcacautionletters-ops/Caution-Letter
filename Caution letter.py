import streamlit as st
import pandas as pd
from fpdf import FPDF

def generate_pdf(df):
    # Use 'UTF-8' friendly settings
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Label Specs: 10cm x 4.3cm
    label_w = 100
    label_h = 43
    
    # Margins: Top 1cm, Right 0.7cm
    top_margin = 10    
    left_margin = 3 # (210mm - 200mm - 7mm right margin = 3mm left)
    
    for i, row in df.iterrows():
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
        
        # --- CLEANING DATA (Crucial to avoid Unicode Errors) ---
        def clean_text(text):
            if pd.isna(text): return ""
            # Replace common problematic characters
            return str(text).encode('ascii', 'ignore').decode('ascii')

        father = clean_text(row['FATHER'])
        name = clean_text(row['NAME'])
        address = clean_text(row['ADDRESS'])
        roll = clean_text(row['ROLL NUMBER'])
        
        ph_parent = clean_text(row['PARENT PHONE'])
        ph_student = clean_text(row['STUDENT PHONE'])
        contact = f"{ph_student}/{ph_parent}".strip("/")
        
        # From Header (Bold)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Body
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 5)
        
        label_text = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {name}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        pdf.multi_cell(label_w - 10, 5, label_text)
        
    # Return output as bytes directly
    return pdf.output(dest='S')

# ... (rest of your file upload logic) ...

if file_caution and file_master:
    # ... (loading data) ...
    
    if st.button("Generate 12-Label PDF"):
        if not df_final.empty:
            try:
                pdf_output = generate_pdf(df_final)
                # Note: fpdf's output('S') returns bytes in newer versions
                # or a string we can convert. This is the safest way:
                st.download_button(
                    label="Download PDF",
                    data=bytes(pdf_output),
                    file_name="Labels.pdf",
                    mime="application/pdf"
                )
                st.success(f"Matched {len(df_final)} students.")
            except Exception as e:
                st.error(f"Error generating PDF: {e}")
