import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --- 1. Helper Functions ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    # Clean float artifacts like '99.0' -> '99'
    return str(val).replace('.0', '').strip()

def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6

# --- 2. Custom PDF Class ---
class LabelPDF(FPDF):
    def __init__(self):
        # We use 'mm' units for physical precision
        super().__init__(orientation='P', unit='mm', format='A4')
        self.set_auto_page_break(auto=False)
        self.set_margins(0, 0, 0)

# --- 3. Streamlit Interface ---
st.set_page_config(page_title="NovaJet 12-Label Pro", layout="wide")
st.title("🏷️ Precision PDF Label Generator (NovaJet 12L)")

# Sidebar for Calibration
st.sidebar.header("🔧 Sheet Calibration")
st.sidebar.info("Adjust these if the print is slightly off on your specific printer.")

# Based on your input: 10mm top, 44mm height, 3mm gap
t_margin = st.sidebar.slider("Top Margin (mm)", 0.0, 20.0, 10.0)
s_margin = st.sidebar.slider("Side Margin (mm)", 0.0, 20.0, 5.0)
l_height = st.sidebar.slider("Label Height (mm)", 30.0, 50.0, 44.0)
l_width = st.sidebar.slider("Label Width (mm)", 80.0, 110.0, 100.0)
v_gap = st.sidebar.slider("Vertical Gap (mm)", 0.0, 10.0, 3.0)

from_addr = st.sidebar.text_area(
    "From Address:", 
    "Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# File Uploads
col1, col2 = st.columns(2)
with col1:
    file_att = st.file_uploader("Upload Attendance Report", type=['xlsx', 'csv'])
    skip = st.number_input("Data starts on row:", min_value=1, value=4)
with col2:
    file_mast = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_att and file_mast:
    try:
        # Load and Match Data
        df_c = pd.read_csv(file_att, skiprows=skip-1) if file_att.name.endswith('csv') else pd.read_excel(file_att, skiprows=skip-1)
        df_m = pd.read_csv(file_mast) if file_mast.name.endswith('csv') else pd.read_excel(file_mast)

        # Cleaning and Filtering
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().unique()
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("🚀 Generate PDF Labels"):
            # Sorting
            df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
            df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])
            records = df_matched.to_dict('records')

            pdf = LabelPDF()
            pdf.set_font("Arial", size=8)

            idx = 0
            while idx < len(records):
                pdf.add_page()
                for row in range(6): # 6 Labels down
                    for col in range(2): # 2 Labels across
                        if idx >= len(records): break
                        
                        # Calculate XY coordinates
                        x = s_margin + (col * l_width)
                        y = t_margin + (row * (l_height + v_gap))
                        
                        # Add Label Content
                        d = records[idx]
                        f_ph = clean_val(d.get('Father_Phone'))
                        s_ph = clean_val(d.get('Student_Phone'))
                        contact = f"{f_ph} / {s_ph}".strip(" / ")
                        
                        # Drawing invisible box for text containment
                        pdf.set_xy(x + 3, y + 5) # Internal padding
                        content = (
                            f"From: {from_addr}\n\n"
                            f"To, Shri/Smt. {clean_val(d.get('Father'))}\n"
                            f"c/o: {clean_val(d.get('Name'))}\n"
                            f"Address: {clean_val(d.get('Address'))}\n"
                            f"Contact: {contact}\n"
                            f"ID: {clean_val(d.get('Roll_No'))}"
                        )
                        pdf.multi_cell(l_width - 6, 4, content, border=0, align='L')
                        idx += 1

            # Output
            pdf_bytes = pdf.output()
            st.success(f"Successfully generated {len(records)} labels!")
            st.download_button(
                label="📥 Download PDF",
                data=pdf_bytes,
                file_name="Student_Labels_A4.pdf",
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"An error occurred: {e}")
