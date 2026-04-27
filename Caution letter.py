import streamlit as st
import pandas as pd
import io

# --- Helper Function: Clean Phone Numbers & IDs ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    return str(val).replace('.0', '').strip()

# --- Custom Sorting Function ---
def get_sort_rank(roll):
    roll = str(roll).upper()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Precision A4 Version)")

# --- SIDEBAR SETTINGS ---
st.sidebar.header("Global Settings")
from_address = st.sidebar.text_area(
    "Edit 'From' Address:", 
    value="Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Attendance Report")
    file_caution = st.file_uploader("Upload Attendance/Shortage File", type=['xlsx', 'csv'], key="caution")
    skip_rows = st.number_input("Attendance data starts on row:", min_value=1, value=4)

with col2:
    st.subheader("2. Master Database")
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'], key="master")

if file_caution and file_master:
    try:
        # Load Files
        df_c = pd.read_csv(file_caution, skiprows=skip_rows-1) if file_caution.name.endswith('csv') else pd.read_excel(file_caution, skiprows=skip_rows-1)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        # Match and Filter
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().unique()
        mast_data = df_m.iloc[:, [1, 5, 29, 18, 45, 44]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        mast_data['Roll_No'] = mast_data['Roll_No'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_matched = mast_data[mast_data['Roll_No'].isin(caution_rolls)].copy()

        if st.button("Generate Precision A4 Labels"):
            if not df_matched.empty:
                df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])

                output = io.BytesIO()
                workbook = pd.ExcelWriter(output, engine='xlsxwriter')
                ws = workbook.book.add_worksheet('PRINT_LABELS')

                # --- 1. COLUMN WIDTHS (100mm = ~46.5 in Excel width units) ---
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.0)  # Tiny horizontal buffer if needed
                ws.set_column('C:C', 46.5)

                label_format = workbook.book.add_format({
                    'font_name': 'Calibri',
                    'font_size': 10,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'align': 'left',
                    'border': 1,
                    'border_color': '#D3D3D3' 
                })

                # --- 2. THE LOOP (Handling 3mm vertical gaps) ---
                r_num = 0 
                data_list = df_matched.to_dict('records')
                
                for i in range(0, len(data_list), 2):
                    # Set height for the label row (44mm = 124.7 points)
                    ws.set_row(r_num, 124.7)

                    # Left Label
                    d = data_list[i]
                    contact = f"{clean_val(d.get('Father_Phone'))} / {clean_val(d.get('Student_Phone'))}".strip(" / ")
                    txt_l = (f"From: {from_address}\n"
                             f"To, Shri/Smt. {clean_val(d.get('Father'))}\n"
                             f"c/o: {clean_val(d.get('Name'))}\n"
                             f"Address: {clean_val(d.get('Address'))}\n"
                             f"Contact: {contact}   ID: {clean_val(d.get('Roll_No'))}")
                    ws.write(r_num, 0, txt_l, label_format)

                    # Right Label
                    if i + 1 < len(data_list):
                        dr = data_list[i+1]
                        contact_r = f"{clean_val(dr.get('Father_Phone'))} / {clean_val(dr.get('Student_Phone'))}".strip(" / ")
                        txt_r = (f"From: {from_address}\n"
                                 f"To, Shri/Smt. {clean_val(dr.get('Father'))}\n"
                                 f"c/o: {clean_val(dr.get('Name'))}\n"
                                 f"Address: {clean_val(dr.get('Address'))}\n"
                                 f"Contact: {contact_r}   ID: {clean_val(dr.get('Roll_No'))}")
                        ws.write(r_num, 2, txt_r, label_format)

                    r_num += 1
                    
                    # --- ADD 3MM GAP ROW ---
                    # Only add gap if we aren't at the very last row of the page
                    # 3mm = ~8.5 points
                    ws.set_row(r_num, 8.5)
                    r_num += 1

                # --- 3. PAGE SETUP (The Alignment "Anchor") ---
                ws.set_paper(9)  # Force A4
                # Top margin = 1.0cm (0.39 inches), Left/Right = 0.5cm (0.2 inches)
                ws.set_margins(left=0.2, right=0.2, top=0.39, bottom=0.2)
                ws.set_print_scale(100) # CRITICAL: Prevents Excel from shrinking the page
                ws.center_horizontally()

                workbook.close()
                st.success("Labels Generated!")
                st.download_button("📥 Download Labels", output.getvalue(), "NovaJet_12_Labels.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
