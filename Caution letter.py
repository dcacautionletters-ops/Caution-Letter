import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --- Helper Function: Clean Data ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    return str(val).replace('.0', '').strip()

def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6

# --- PDF Class Definition ---
class LabelPDF(FPDF):
    def __init__(self):
        super().__init__(orientation='P', unit='mm', format='A4')
        self.set_auto_page_break(auto=False)

st.set_page_config(page_title="PDF Label Pro", layout="wide")
st.title("🏷️ Precision PDF Label Generator")
st.info("This version generates a locked PDF to prevent 'Bottom Drift'.")

# --- SIDEBAR ---
st.sidebar.header("Label Dimensions (mm)")
top_margin = st.sidebar.number_input("Top Margin (mm)", value=10.0)
left_margin = st.sidebar.number_input("Side Margin (mm)", value=5.0)
label_h = st.sidebar.number_input("Label Height (mm)", value=44.0)
label_w = st.sidebar.number_input("Label Width (mm)", value=100.0)
v_gap = st.sidebar.number_input("Vertical Gap (mm)", value=3.0)
h_gap = st.sidebar.number_input("Horizontal Gap (mm)", value=0.0)

from_addr = st.sidebar.text_area("From Address:", "Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024")

# --- FILE UPLOADS ---
col1, col2 = st.columns(2)
with col1:
    file_caution = st.file_uploader("Upload Attendance File", type=['xlsx', 'csv'])
    skip_rows = st.number_input("Data starts on row:", min_value=1, value=4)
with col2:
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_caution and file_master:
    try:
        df_c = pd.read_csv(file_caution, skiprows=skip_rows-1) if file_caution.name.endswith('csv') else pd.read_excel(file_caution, skiprows=skip_rows-1)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().unique()
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("Generate PDF Labels"):
            df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
            df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])
            records = df_matched.to_dict('records')

            pdf = LabelPDF()
            pdf.set_font("Arial", size=8)

            # Start coordinates
            x_start = left_margin
            y_start = top_margin
            
            curr_rec = 0
            while curr_rec < len(records):
                pdf.add_page()
                # 6 rows per page
                for row in range(6):
                    # 2 columns per row
                    for col in range(2):
                        if curr_rec >= len(records): break
                        
                        # Calculate position
                        x = x_start + (col * (label_w + h_gap))
                        y = y_start + (row * (label_h + v_gap))
                        
                        # Draw Label Border (Optional, helpful for testing)
                        pdf.set_draw_color(200, 200, 200)
                        pdf.rect(x, y, label_w, label_h)
                        
                        # Fill Content
                        d = records[curr_rec]
                        contact = f"{clean_val(d.get('Father_Phone'))} / {clean_val(d.get('Student_Phone'))}".strip(" / ")
                        
                        pdf.set_xy(x + 2, y + 2) # Padding inside label
                        pdf.multi_cell(label_w - 4, 4, 
                            f"From: {from_addr}\n"
                            f"To, Shri/Smt. {clean_val(d.get('Father'))}\n"
                            f"c/o: {clean_val(d.get('Name'))}\n"
                            f"Address: {clean_val(d.get('Address'))}\n"
                            f"Contact: {contact}   ID: {clean_val(d.get('Roll_No'))}", 
                            border=0, align='L')
                        
                        curr_rec += 1

            pdf_output = pdf.output()
            st.success(f"Created {len(records)} labels.")
            st.download_button("📥 Download PDF for Printing", data=bytes(pdf_output), file_name="Labels_Final.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"Error: {e}")
