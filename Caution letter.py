import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --- 1. Helper Functions ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    # Strip decimals like 123.0 -> 123
    return str(val).split('.')[0].strip()

def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6

# --- 2. Interface ---
st.set_page_config(page_title="NovaJet Label Pro", layout="wide")
st.title("🏷️ Precision Label Generator (Final Stabilized)")

# Sidebar Settings
st.sidebar.header("Calibration")
t_margin = st.sidebar.slider("Top Margin (mm)", 0.0, 25.0, 10.0)
s_margin = st.sidebar.slider("Side Margin (mm)", 0.0, 20.0, 5.0)
v_gap = st.sidebar.slider("Vertical Gap (mm)", 0.0, 10.0, 3.0)

from_addr = st.sidebar.text_area(
    "From Address:", 
    "Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# File Upload Section
col1, col2 = st.columns(2)
with col1:
    file_att = st.file_uploader("Upload Attendance Report", type=['xlsx', 'csv'])
    skip_row_val = st.number_input("Data starts on row:", min_value=1, value=4)

with col2:
    file_mast = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_att and file_mast:
    try:
        # Load Attendance
        df_c = pd.read_csv(file_att, skiprows=skip_row_val-1) if file_att.name.endswith('csv') else pd.read_excel(file_att, skiprows=skip_row_val-1)
            
        # Load Master
        df_m = pd.read_csv(file_mast) if file_mast.name.endswith('csv') else pd.read_excel(file_mast)

        # Match and Clean
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.split('.').str[0].str.strip().unique()
        
        # Select Columns: Roll(B), Name(F), Father(AD), Address(S), FatherPh(AT), StudentPh(AS)
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.split('.').str[0].str.strip()
        
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("🚀 Generate PDF Labels"):
            if not df_matched.empty:
                # Sort the data
                df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])
                records = df_matched.to_dict('records')

                # Create PDF in memory
                pdf = FPDF(orientation='P', unit='mm', format='A4')
                pdf.set_auto_page_break(auto=False)
                pdf.set_font("Helvetica", size=9)

                idx = 0
                while idx < len(records):
                    pdf.add_page()
                    for row in range(6): 
                        for col in range(2): 
                            if idx >= len(records): break
                            
                            # Position (100x44 labels)
                            x = s_margin + (col * 100)
                            y = t_margin + (row * (44 + v_gap))
                            
                            d = records[idx]
                            contact = f"{clean_val(d.get('Father_Phone'))} / {clean_val(d.get('Student_Phone'))}".strip(" / ")
                            
                            pdf.set_xy(x + 3, y + 5)
                            content = (
                                f"From: {from_addr}\n"
                                f"To: Shri/Smt. {clean_val(d.get('Father'))}\n"
                                f"c/o: {clean_val(d.get('Name'))}\n"
                                f"Address: {clean_val(d.get('Address'))}\n"
                                f"Contact: {contact}   ID: {clean_val(d.get('Roll_No'))}"
                            )
                            pdf.multi_cell(94, 4.5, content, border=0)
                            idx += 1

                # --- THE BUFFER METHOD ---
                # We write the PDF to a string, then convert to bytes
                # This ensures Adobe sees a perfectly clean binary header
                pdf_output = pdf.output()
                
                # Using a manual byte conversion to ensure total Adobe compatibility
                if isinstance(pdf_output, str):
                    final_pdf_bytes = pdf_output.encode('latin-1')
                else:
                    final_pdf_bytes = bytes(pdf_output)

                st.success(f"Generated {len(records)} labels.")
                st.download_button(
                    label="📥 Download & Open in Adobe",
                    data=final_pdf_bytes,
                    file_name="Labels_For_Adobe.pdf",
                    mime="application/pdf"
                )
            else:
                st.warning("No matches found between the files.")

    except Exception as e:
        st.error(f"An error occurred: {e}")
