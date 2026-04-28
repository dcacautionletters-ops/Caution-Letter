import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --- Helper Functions ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    return str(val).split('.')[0].strip()

def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6

st.set_page_config(page_title="NovaJet Label Pro", layout="wide")
st.title("🏷️ Student Label Generator (Final Fix)")

# Sidebar Settings
st.sidebar.header("Calibration")
t_margin = st.sidebar.slider("Top Margin (mm)", 0.0, 20.0, 10.0)
s_margin = st.sidebar.slider("Side Margin (mm)", 0.0, 20.0, 5.0)
v_gap = st.sidebar.slider("Vertical Gap (mm)", 0.0, 10.0, 3.0)

from_addr = st.sidebar.text_area("From:", "Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024")

# File Upload
col1, col2 = st.columns(2)
with col1:
    file_att = st.file_uploader("Attendance File", type=['xlsx', 'csv'])
    skip_val = st.number_input("Start Row", min_value=1, value=4)
with col2:
    file_mast = st.file_uploader("Master File", type=['xlsx', 'csv'])

if file_att and file_mast:
    try:
        # Load Data
        df_c = pd.read_csv(file_att, skiprows=skip_val-1) if file_att.name.endswith('csv') else pd.read_excel(file_att, skiprows=skip_row_val-1)
        df_m = pd.read_csv(file_mast) if file_mast.name.endswith('csv') else pd.read_excel(file_mast)

        # Match Logic
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.split('.').str[0].str.strip().unique()
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.split('.').str[0].str.strip()
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("🚀 Generate PDF"):
            df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
            df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])
            records = df_matched.to_dict('records')

            # PDF Setup
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=False)
            pdf.set_font("Helvetica", size=9)

            idx = 0
            while idx < len(records):
                pdf.add_page()
                for row in range(6):
                    for col in range(2):
                        if idx >= len(records): break
                        
                        # Position Math (100x44mm labels)
                        x = s_margin + (col * 100)
                        y = t_margin + (row * (44 + v_gap))
                        
                        d = records[idx]
                        contact = f"{clean_val(d.get('Father_Phone'))} / {clean_val(d.get('Student_Phone'))}".strip(" / ")
                        
                        pdf.set_xy(x + 2, y + 5)
                        text = (f"From: {from_addr}\nTo: {clean_val(d.get('Father'))}\n"
                                f"c/o: {clean_val(d.get('Name'))}\nAddress: {clean_val(d.get('Address'))}\n"
                                f"Contact: {contact}  ID: {clean_val(d.get('Roll_No'))}")
                        pdf.multi_cell(96, 4, text)
                        idx += 1

            # --- THE FAIL-SAFE OUTPUT METHOD ---
            # We use a memory buffer to ensure NO characters are added to the PDF binary
            pdf_output = pdf.output() 
            
            # If pdf.output() returns a string (older versions), we encode it.
            # If it's already a bytearray/bytes, we leave it alone.
            if isinstance(pdf_output, str):
                final_data = pdf_output.encode('latin-1')
            else:
                final_data = bytes(pdf_output)

            st.success("PDF Ready!")
            st.download_button(
                label="📥 Download & Open in Adobe",
                data=final_data,
                file_name="Student_Labels_Final.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error: {e}")
