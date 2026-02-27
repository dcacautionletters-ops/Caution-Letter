import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

def generate_labels(df_final):
    # PDF Settings: A4 (210x297mm)
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_page()
    
    # Precise Spacing from your requirements
    margin_top = 10    # 1 cm
    margin_left = 7    # 0.7 cm (adjusted for Right/Left balance)
    label_width = 100  # 10 cm
    label_height = 43  # 4.3 cm
    cols = 2
    
    x_start = margin_left
    y_start = margin_top
    
    for i, row in df_final.iterrows():
        # Calculate Grid Position
        col_idx = i % cols
        row_idx = (i // cols) % 6 # 6 rows per page (12 labels)
        
        if i > 0 and i % 12 == 0:
            pdf.add_page()
            
        x = x_start + (col_idx * label_width)
        y = y_start + (row_idx * label_height)
        
        # Draw Border (Light Grey like VBA)
        pdf.set_draw_color(200, 200, 200)
        pdf.rect(x, y, label_width, label_height)
        
        # --- Label Content ---
        pdf.set_xy(x + 2, y + 5)
        
        # From Line (Bold)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 4, "From: Presidency College Bangalore (AUTONOMOUS)", ln=True)
        
        # Body Text (Regular)
        pdf.set_font("Arial", '', 9)
        pdf.set_x(x + 2)
        pdf.multi_cell(label_width - 4, 5, 
            f"To, Shri/Smt. {row['Father']}\n" +
            f"c/o: {row['StudentName']}\n" +
            f"Address: {row['Address']}\n" +
            f"Contact: {row['Phones']}   ID: {row['RollNo']}"
        )

    return pdf.output(dest='S').encode('latin-1')

# --- Streamlit UI ---
st.title("🏷️ Universal Label Generator")
st.write("Upload your files to generate the 2x6 Precision Labels.")

col1, col2 = st.columns(2)

with col1:
    caution_file = st.file_uploader("Upload Caution Letter List", type=['xlsx', 'csv'])
with col2:
    master_file = st.file_uploader("Upload Student Master Database", type=['xlsx', 'csv'])

if caution_file and master_file:
    # Load Data
    df_caution = pd.read_csv(caution_file) if caution_file.name.endswith('csv') else pd.read_excel(caution_file)
    df_master = pd.read_csv(master_file) if master_file.name.endswith('csv') else pd.read_excel(master_file)
    
    try:
        # 1. Clean data (Trim spaces)
        df_caution['RollNo'] = df_caution.iloc[:, 1].astype(str).str.strip() # Col B
        df_master['RollNo'] = df_master.iloc[:, 1].astype(str).str.strip()  # Col B
        
        # 2. Merge (Core Logic: Exactly match Roll No and Student Name)
        # Caution: Roll(B), Name(C) | Master: Roll(B), Name(F)
        df_merged = pd.merge(
            df_caution, 
            df_master, 
            left_on=[df_caution.columns[1], df_caution.columns[2]], 
            right_on=[df_master.columns[1], df_master.columns[5]], 
            how='inner'
        )
        
        # 3. Extract required columns
        final_data = []
        for _, row in df_merged.iterrows():
            # Join phone numbers with /
            ph1 = str(row.iloc[43]) if not pd.isna(row.iloc[43]) else "" # AR (43)
            ph2 = str(row.iloc[44]) if not pd.isna(row.iloc[44]) else "" # AS (44)
            phones = f"{ph1}/{ph2}".strip("/")
            
            final_data.append({
                "RollNo": row.iloc[1],      # Col B
                "StudentName": row.iloc[5], # Col F
                "Father": row.iloc[28],      # Col AC (Index 28)
                "Address": row.iloc[17],     # Col R (Index 17)
                "Phones": phones
            })
        
        df_final = pd.DataFrame(final_data)
        
        if st.button("Generate Labels PDF"):
            pdf_bytes = generate_labels(df_final)
            st.download_button(
                label="📥 Download Labels for Printing",
                data=pdf_bytes,
                file_name="Student_Labels.pdf",
                mime="application/pdf"
            )
            st.success(f"Processed {len(df_final)} students.")
            
    except Exception as e:
        st.error(f"Error matching columns: {e}")