import streamlit as st
import pandas as pd
from fpdf import FPDF

# --- PDF GENERATION LOGIC ---
def generate_pdf(df):
    # Standard A4: 210 x 297 mm
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False) # Manual control for precision
    pdf.add_page()
    
    # STRICT DIMENSIONS (mm)
    label_w = 100   # 10 cm
    label_h = 43    # 4.3 cm
    
    # STRICT MARGINS (mm)
    top_margin = 10    # 1 cm from top
    left_margin = 7    # 0.7 cm from left
    
    for i, (_, row) in enumerate(df.iterrows()):
        # Calculate grid coordinates
        col = i % 2
        line = (i // 2) % 6
        
        # Add new page every 12 labels
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        x = left_margin + (col * label_w)
        y = top_margin + (line * label_h)
        
        # 1. Draw the Label Box (Light Grey)
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        # 2. Text Content (Ensuring it stays INSIDE the box)
        # Margin inside the box = 4mm
        text_x = x + 4
        text_y = y + 5
        pdf.set_xy(text_x, text_y)
        
        # Helper to clean text
        def clean(val):
            if pd.isna(val): return ""
            return str(val).encode('ascii', 'ignore').decode('ascii')

        # Data mapping
        father = clean(row.get('FATHER', ''))
        student_name = clean(row.get('NAME', ''))
        address = clean(row.get('ADDRESS', ''))
        roll = clean(row.get('ROLL NUMBER', ''))
        ph_s = clean(row.get('STUDENT PHONE', ''))
        ph_p = clean(row.get('PARENT PHONE', ''))
        contact = f"{ph_s}/{ph_p}".strip("/")
        
        # Header (Bold)
        pdf.set_font("Arial", 'B', 8.5) # Slightly smaller font to ensure fit
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Body (Regular)
        pdf.set_font("Arial", '', 8.5)
        pdf.set_x(text_x)
        
        label_text = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {student_name}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        # multi_cell width set to label_w - 8 to keep it away from the borders
        pdf.multi_cell(label_w - 8, 4.5, label_text)
        
    # Generate Output
    pdf_out = pdf.output(dest='S')
    if isinstance(pdf_out, str):
        return pdf_out.encode('latin-1')
    return pdf_out

# --- UI LOGIC ---
st.title("Precise Address Label Generator")

c_file = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'])
m_file = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if c_file and m_file:
    try:
        df_c = pd.read_csv(c_file) if c_file.name.endswith('csv') else pd.read_excel(c_file)
        df_m = pd.read_csv(m_file) if m_file.name.endswith('csv') else pd.read_excel(m_file)

        # Match by Roll Number (Column B in Caution, ROLL NUMBER in Master)
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.strip()
        
        df_final = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]

        if st.button("Generate 12-Label PDF"):
            if not df_final.empty:
                pdf_bytes = generate_pdf(df_final)
                st.download_button(
                    label="📥 Download Labels PDF",
                    data=pdf_bytes,
                    file_name="Student_Labels_A4.pdf",
                    mime="application/pdf"
                )
                st.success(f"Generated labels for {len(df_final)} students.")
            else:
                st.error("No matches found between the files.")
    except Exception as e:
        st.error(f"Error: {e}")
