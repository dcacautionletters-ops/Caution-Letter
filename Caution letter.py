import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --- 1. Helper Functions ---
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

# --- 2. Main Streamlit App ---
st.set_page_config(page_title="NovaJet PDF Labeler", layout="wide")
st.title("🏷️ Precision Label Generator")

# Sidebar Calibration
st.sidebar.header("📏 Calibration Settings")
t_margin = st.sidebar.slider("Top Margin (mm)", 0.0, 20.0, 10.0)
s_margin = st.sidebar.slider("Side Margin (mm)", 0.0, 20.0, 5.0)
l_height = st.sidebar.slider("Label Height (mm)", 30.0, 50.0, 44.0)
l_width = st.sidebar.slider("Label Width (mm)", 80.0, 110.0, 100.0)
v_gap = st.sidebar.slider("Vertical Gap (mm)", 0.0, 10.0, 3.0)

from_addr = st.sidebar.text_area("From Address:", "Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024")

# File Upload Section
col1, col2 = st.columns(2)
with col1:
    file_att = st.file_uploader("Upload Attendance File", type=['xlsx', 'csv'])
    skip_row_val = st.number_input("Data starts on row:", min_value=1, value=4)
with col2:
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_att and file_master:
    try:
        # Load Data
        df_c = pd.read_csv(file_att, skiprows=skip_row_val-1) if file_att.name.endswith('csv') else pd.read_excel(file_att, skiprows=skip_row_val-1)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        # Match Data
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.split('.').str[0].str.strip().unique()
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.split('.').str[0].str.strip()
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("🚀 Generate PDF for Printing"):
            df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
            df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])
            records = df_matched.to_dict('records')

            # PDF Generation
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=False)
            pdf.set_font("Helvetica", size=9)

            idx = 0
            while idx < len(records):
                pdf.add_page()
                for row in range(6):
                    for col in range(2):
                        if idx >= len(records): break
                        
                        x = s_margin + (col * l_width)
                        y = t_margin + (row * (l_height + v_gap))
                        
                        d = records[idx]
                        contact = f"{clean_val(d.get('Father_Phone'))} / {clean_val(d.get('Student_Phone'))}".strip(" / ")
                        
                        pdf.set_xy(x + 2, y + 5)
                        label_content = (
                            f"From: {from_addr}\n"
                            f"To, Shri/Smt. {clean_val(d.get('Father'))}\n"
                            f"c/o: {clean_val(d.get('Name'))}\n"
                            f"Address: {clean_val(d.get('Address'))}\n"
                            f"Contact: {contact}   ID: {clean_val(d.get('Roll_No'))}"
                        )
                        pdf.multi_cell(l_width - 4, 4, label_content, border=0)
                        idx += 1

            # CRITICAL FIX FOR ADOBE ERROR:
            # Convert the bytearray to bytes explicitly
            pdf_output = pdf.output()
            if isinstance(pdf_output, bytearray):
                pdf_bytes = bytes(pdf_output)
            else:
                pdf_bytes = pdf_output

            st.success(f"Generated {len(records)} labels.")
            st.download_button(
                label="📥 Download PDF for Adobe",
                data=pdf_bytes,
                file_name="Student_Labels.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error: {e}")
