import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --- PDF GENERATION LOGIC ---
def generate_pdf(df):
    # Standard A4: 210 x 297 mm
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Label Dimensions (10cm x 4.3cm)
    label_w = 100
    label_h = 43
    
    # Precise Margins (Top 1cm, Right 0.7cm)
    top_margin = 10    
    right_margin = 7
    # Calculate Left Margin to balance 0.7cm right on 210mm page
    left_margin = 210 - (label_w * 2) - right_margin # Result: 3mm
    
    for i, (_, row) in enumerate(df.iterrows()):
        col = i % 2
        line = (i // 2) % 6
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        x = left_margin + (col * label_w)
        y = top_margin + (line * label_h)
        
        # Draw Border
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        # Text Offsets
        pdf.set_xy(x + 5, y + 5)
        
        # Clean data to avoid encoding crashes
        def clean(val):
            if pd.isna(val): return ""
            return str(val).encode('ascii', 'ignore').decode('ascii')

        # Map exactly to your StudentList-details.xlsx headers
        father = clean(row.get('FATHER', ''))
        student_name = clean(row.get('NAME', ''))
        address = clean(row.get('ADDRESS', ''))
        roll = clean(row.get('ROLL NUMBER', ''))
        ph_student = clean(row.get('STUDENT PHONE', ''))
        ph_parent = clean(row.get('PARENT PHONE', ''))
        
        contact = f"{ph_student}/{ph_parent}".strip("/")
        
        # Label Content
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 5)
        
        # Formatting the "To" and "c/o" lines
        label_text = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {student_name}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        pdf.multi_cell(label_w - 10, 5, label_text)
        
    # FIX: Return as a byte string without manual encoding
    return pdf.output(dest='S')

# --- STREAMLIT UI ---
st.title("🏷️ Student Label Generator")

file_caution = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'])
file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_caution and file_master:
    try:
        # Load Files
        df_c = pd.read_csv(file_caution) if file_caution.name.endswith('csv') else pd.read_excel(file_caution)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        # 1. Get Roll Numbers from Caution Letter (Col B / Index 1)
        # We strip spaces to ensure a perfect match
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        
        # 2. Filter Master Data where ROLL NUMBER matches
        # Ensure 'ROLL NUMBER' column is treated as a string
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.strip()
        df_final = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]

        if st.button("Generate Labels"):
            if not df_final.empty:
                # Get the PDF as a string or bytes
                pdf_data = generate_pdf(df_final)
                
                # Convert to bytes if it's currently a string (handles library version differences)
                if isinstance(pdf_data, str):
                    pdf_data = pdf_data.encode('latin-1')

                st.download_button(
                    label="📥 Download 12-Label PDF",
                    data=pdf_data,
                    file_name="Caution_Labels.pdf",
                    mime="application/pdf"
                )
                st.success(f"Matched {len(df_final)} students.")
            else:
                st.error("No matches found. Check if Roll Numbers in both files are identical.")

    except Exception as e:
        st.error(f"Error processing files: {e}")
