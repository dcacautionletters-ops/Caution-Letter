import streamlit as st
import pandas as pd
from fpdf import FPDF

# --- PDF GENERATION LOGIC ---
def generate_pdf(df):
    # Standard A4: 210 x 297 mm
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    
    # --- TWEAK: REMOVE DEFAULT MARGINS ---
    # We set margins to 0 so our manual X, Y coordinates are absolute
    pdf.set_margins(left=0, top=0, right=0)
    pdf.set_auto_page_break(auto=False) 
    
    pdf.add_page()
    
    # STRICT DIMENSIONS (mm)
    label_w = 100   # 10 cm
    label_h = 43    # 4.3 cm
    
    # STRICT MARGINS FROM PAGE EDGE (mm)
    top_offset = 10    # 1 cm from top edge of paper
    left_offset = 7    # 0.7 cm from left edge of paper
    
    for i, (_, row) in enumerate(df.iterrows()):
        col = i % 2
        line = (i // 2) % 6
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        # Calculate absolute X and Y
        x = left_offset + (col * label_w)
        y = top_offset + (line * label_h)
        
        # Draw the Label Box
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        # Position text with a very small padding inside the box (2mm)
        pdf.set_xy(x + 2, y + 3)
        
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
        
        # From Header (Bold)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 4, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Body (Regular)
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 2)
        
        label_text = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {student_name}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        # multi_cell width allows text to fill the 100mm width
        pdf.multi_cell(label_w - 4, 5, label_text)
        
    return pdf.output(dest='S')

# --- UI LOGIC ---
st.title("Precise Address Label Generator")

c_file = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'])
m_file = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if c_file and m_file:
    try:
        df_c = pd.read_csv(c_file) if c_file.name.endswith('csv') else pd.read_excel(c_file)
        df_m = pd.read_csv(m_file) if m_file.name.endswith('csv') else pd.read_excel(m_file)

        # Match logic
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.strip()
        df_final = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]

        if st.button("Generate 12-Label PDF"):
            if not df_final.empty:
                pdf_output = generate_pdf(df_final)
                
                # Check for version compatibility (string vs bytes)
                pdf_bytes = pdf_output.encode('latin-1') if isinstance(pdf_output, str) else pdf_bytes = pdf_output
                
                st.download_button(
                    label="📥 Download Labels PDF",
                    data=pdf_bytes,
                    file_name="Student_Labels_Final.pdf",
                    mime="application/pdf"
                )
                st.success(f"Generated {len(df_final)} labels.")
    except Exception as e:
        st.error(f"Error: {e}")
