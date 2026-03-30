import streamlit as st
import pandas as pd
import io

# --- Helper Function: Clean Phone Numbers & IDs ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    text = str(val).replace('.0', '').strip()
    return text

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
st.title("🏷️ Student Label Generator (Rectified)")

# --- SIDEBAR ---
st.sidebar.header("Global Settings")
from_address = st.sidebar.text_area(
    "Edit 'From' Address:", 
    value="Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# --- FILE UPLOAD ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Shortage Report")
    file_shortage = st.file_uploader("Upload Shortage File", type=['xlsx', 'csv'], key="shortage")
    skip_rows = st.number_input("Shortage data starts on row (usually 4):", min_value=1, value=4)

with col2:
    st.subheader("2. Master Database")
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'], key="master")

if file_shortage and file_master:
    try:
        # 1. Load Shortage File
        if file_shortage.name.endswith('csv'):
            df_s = pd.read_csv(file_shortage, skiprows=skip_rows-1)
        else:
            df_s = pd.read_excel(file_shortage, skiprows=skip_rows-1)
        
        # Extract unique Roll Numbers from Shortage Column B (Index 1)
        shortage_rolls = df_s.iloc[:, 1].dropna().astype(str).str.replace('.0', '', regex=False).str.strip().str.upper().unique()
        
        # 2. Load Master File
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)

        # Mapping Master Columns (Indices): 
        # B=1, F=5, J=9, AE=30, T=19, AU=46, AT=45
        mast_data = df_m.iloc[:, [1, 5, 30, 19, 46, 45]].copy()
        mast_data.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Phone_AU', 'Phone_AT']
        
        # Clean Master Roll Nos for matching
        mast_data['Match_ID'] = mast_data['Roll_No'].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()

        # 3. Filter Master Data: Get students whose IDs are in the Shortage Report
        df_matched = mast_data[mast_data['Match_ID'].isin(shortage_rolls)].copy()

        # 4. Remove Duplicates: 1 Label per Roll Number
        df_matched = df_matched.drop_duplicates(subset=['Match_ID'])

        # Display Summary for Transparency
        st.write(f"📊 **Summary:** Found {len(shortage_rolls)} unique students in Shortage Report. Matched {len(df_matched)} students in Master Database.")

        if st.button("Generate Sorted A4 Label Sheet"):
            if not df_matched.empty:
                # --- SORTING ---
                df_matched['sort_rank'] = df_matched['Match_ID'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'Match_ID'])

                # Create Excel Workbook
                output = io.BytesIO()
                workbook_writer = pd.ExcelWriter(output, engine='xlsxwriter')
                ws = workbook_writer.book.add_worksheet('PRINT_LABELS')

                # --- FORMATTING (A4 Alignment) ---
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.5) # Gap
                ws.set_column('C:C', 46.5)

                label_format = workbook_writer.book.add_format({
                    'font_name': 'Calibri',
                    'font_size': 10,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'border': 1,
                    'border_color': '#C8C8C8' 
                })

                # --- LOOP THROUGH DATA (2 per row) ---
                r_num = 0 
                data_list = df_matched.to_dict('records')
                num_records = len(data_list)
                
                for i in range(0, num_records, 2):
                    ws.set_row(r_num, 122) # Label Height

                    # --- Left Label (Col A) ---
                    d = data_list[i]
                    p1, p2 = clean_val(d.get('Phone_AU')), clean_val(d.get('Phone_AT'))
                    contact = f"{p1} / {p2}".strip(" / ")
                    
                    txt_left = (f"From: {from_address}\n"
                                f"To, Shri/Smt. {clean_val(d.get('Father'))}\n"
                                f"c/o: {clean_val(d.get('Name'))}\n"
                                f"Address: {clean_val(d.get('Address'))}\n"
                                f"Contact: {contact}   ID: {clean_val(d.get('Roll_No'))}")
                    ws.write(r_num, 0, txt_left, label_format)

                    # --- Right Label (Col C) ---
                    if i + 1 < num_records:
                        dr = data_list[i+1]
                        pr1, pr2 = clean_val(dr.get('Phone_AU')), clean_val(dr.get('Phone_AT'))
                        contact_r = f"{pr1} / {pr2}".strip(" / ")
                        
                        txt_right = (f"From: {from_address}\n"
                                     f"To, Shri/Smt. {clean_val(dr.get('Father'))}\n"
                                     f"c/o: {clean_val(dr.get('Name'))}\n"
                                     f"Address: {clean_val(dr.get('Address'))}\n"
                                     f"Contact: {contact_r}   ID: {clean_val(dr.get('Roll_No'))}")
                        ws.write(r_num, 2, txt_right, label_format)

                    r_num += 1
                    ws.set_row(r_num, 8) # Vertical Gap
                    r_num += 1

                # PAGE SETUP
                ws.set_paper(9) # A4
                ws.set_margins(left=0.2, right=0.2, top=0.3, bottom=0.3)
                ws.set_print_scale(100)
                ws.center_horizontally()

                workbook_writer.close()
                
                st.success(f"Successfully generated {len(df_matched)} unique labels.")
                st.download_button(
                    label="📥 Download Rectified Label Sheet",
                    data=output.getvalue(),
                    file_name="Student_Labels_Rectified.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No students from the Shortage Report were found in the Master Database. Check the Roll Numbers.")
                
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Upload both files to start.")
