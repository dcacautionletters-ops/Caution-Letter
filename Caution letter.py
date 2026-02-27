import streamlit as st
import pandas as pd
from fpdf import FPDF

# --- PDF GENERATION LOGIC ---
def generate_pdf(df):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Label Specs: 10cm x 4.3cm
    label_w = 100
    label_h = 43
    
    # Margins: Top 1cm, Right 0.7cm, calculated Left 0.3cm
    top_margin = 10    
    left_margin = 3 
    
    for i, row in df.iterrows():
        col = i % 2
        line = (i // 2) % 6
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        x = left_margin + (col * label_w)
        y = top_margin + (line * label_h)
        
        # Draw Label Border (Light Grey)
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        pdf.set_xy(x + 5, y + 5)
        
        # Helper to remove non-Latin characters that crash FPDF
        def clean(text):
            if pd.isna(text): return ""
            return str(text).encode('ascii', 'ignore').decode('ascii')

        # Mapping data from your specific Master File headers
        father = clean(row['FATHER'])
        student_name = clean(row['NAME'])
        address = clean(row['ADDRESS'])
        roll = clean(row['ROLL NUMBER'])
        ph_student = clean(row['STUDENT PHONE'])
        ph_parent = clean(row['PARENT PHONE'])
        
        contact = f"{ph_student}/{ph_parent}".strip("/")
        
        # Header (Bold)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Body (Regular)
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 5)
        
        label_text = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {student_name}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        pdf.multi_cell(label_w - 10, 5, label_text)
        
    return pdf.output(dest='S')

# --- STREAMLIT UI ---
st.title("🏷️ Precision Label Generator")
st.write("Match Caution Letter list with Master Database to print 2x6 labels.")

# Variables assigned here must match the IF statement below
file_caution = st.file_uploader("Upload Caution Letter List (CSV/XLSX)", type=['xlsx', 'csv'])
file_master = st.file_uploader("Upload Student Master Data (CSV/XLSX)", type=['xlsx', 'csv'])

if file_caution and file_master:
    # Load Data
    try:
        if file_caution.name.endswith('csv'):
            df_c = pd.read_csv(file_caution)
        else:
            df_c = pd.read_excel(file_caution)
            
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)

        # CORE LOGIC: Match Roll No (Col B) from Caution Letter 
        # to ROLL NUMBER in Master Data
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        
        # Filter master data
        df_final = df_m[df_m['ROLL NUMBER'].astype(str).str.strip().isin(caution_rolls)]

        if st.button("Generate 12-Label PDF"):
            if not df_final.empty:
                pdf_output = generate_pdf(df_final)
                st.download_button(
                    label="📥 Download Labels PDF",
                    data=bytes(pdf_output),
                    file_name="Student_Labels.pdf",
                    mime="application/pdf"
                )
                st.success(f"Matched {len(df_final)} students found in Caution List.")
            else:
                st.error("No matching Roll Numbers found between the two files.")
                
    except Exception as e:
        st.error(f"Error processing files: {e}")
