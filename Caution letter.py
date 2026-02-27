import streamlit as st
import pandas as pd
from fpdf import FPDF

def generate_pdf(df):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_margins(left=0, top=0, right=0)
    pdf.set_auto_page_break(auto=False) 
    pdf.add_page()
    
    # --- STRICT DIMENSIONS ---
    label_w = 100   # 10 cm
    label_h = 43    # 4.3 cm
    row_gap = 3     # 0.3 cm gap between rows
    
    # --- STRICT MARGINS ---
    top_offset = 10    # 1 cm from top
    left_offset = 7    # 0.7 cm from left
    
    for i, (_, row) in enumerate(df.iterrows()):
        col = i % 2
        line = (i // 2) % 6
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        # Calculate X and Y with the new Row Gap logic
        x = left_offset + (col * label_w)
        # Y = Top Margin + (Label Height * row index) + (Gap * row index)
        y = top_offset + (line * label_h) + (line * row_gap)
        
        # Draw Box
        pdf.set_draw_color(220, 220, 220)
        pdf.rect(x, y, label_w, label_h)
        
        # Text start
        pdf.set_xy(x + 2, y + 2)
        
        def clean(val):
            if pd.isna(val): return ""
            return str(val).encode('ascii', 'ignore').decode('ascii')

        father = clean(row.get('FATHER', ''))
        student_name = clean(row.get('NAME', ''))
        address = clean(row.get('ADDRESS', ''))
        roll = clean(row.get('ROLL NUMBER', ''))
        ph_s = clean(row.get('STUDENT PHONE', ''))
        ph_p = clean(row.get('PARENT PHONE', ''))
        contact = f"{ph_s}/{ph_p}".strip("/")
        
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 5, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 2)
        
        label_text = (
            f"To, Shri/Smt. {father}\n"
            f"c/o: {student_name}\n"
            f"Address: {address}\n"
            f"Contact: {contact}   ID: {roll}"
        )
        pdf.multi_cell(label_w - 4, 5, label_text)
        
    return pdf.output(dest='S')

# --- STREAMLIT UI ---
st.title("Precise Address Label Generator")

file_caution = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'])
file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_caution and file_master:
    try:
        df_c = pd.read_csv(file_caution) if file_caution.name.endswith('csv') else pd.read_excel(file_caution)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.strip()
        df_final = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]

        if st.button("Generate 12-Label PDF"):
            if not df_final.empty:
                pdf_output = generate_pdf(df_final)
                
                # Handling version differences for byte conversion
                if isinstance(pdf_output, str):
                    final_bytes = pdf_output.encode('latin-1')
                else:
                    final_bytes = pdf_output
                
                st.download_button(
                    label="📥 Download Labels PDF",
                    data=final_bytes,
                    file_name="Student_Labels_with_Gaps.pdf",
                    mime="application/pdf"
                )
                st.success(f"Generated labels for {len(df_final)} students.")
            else:
                st.error("No matches found.")
    except Exception as e:
        st.error(f"Error: {e}")
