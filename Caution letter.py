import streamlit as st
import pandas as pd
import io

# --- Helper Function: Clean Phone Numbers & IDs ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    # Remove .0 if it's a number stored as a float
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
st.title("🏷️ Student Label Generator (Updated Matching)")

# --- SIDEBAR SETTINGS ---
st.sidebar.header("Global Settings")
from_address = st.sidebar.text_area(
    "Edit 'From' Address:", 
    value="Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Shortage Report")
    file_shortage = st.file_uploader("Upload Shortage File", type=['xlsx', 'csv'], key="shortage")
    skip_rows = st.number_input("Shortage data starts on row:", min_value=1, value=4)

with col2:
    st.subheader("2. Master Database")
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'], key="master")

if file_shortage and file_master:
    try:
        # 1. Load Shortage File (B, C, G)
        if file_shortage.name.endswith('csv'):
            df_s = pd.read_csv(file_shortage, skiprows=skip_rows-1)
        else:
            df_s = pd.read_excel(file_shortage, skiprows=skip_rows-1)
        
        # Clean Shortage keys for matching (B=1, C=2, G=6)
        shortage_keys = df_s.iloc[:, [1, 2, 6]].dropna(subset=[df_s.columns[1]]).copy()
        shortage_keys.columns = ['Match_B', 'Match_C', 'Match_G']
        for col in shortage_keys.columns:
            shortage_keys[col] = shortage_keys[col].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
        
        # 2. Load Master File (B, F, J, AE, T, AU, AT)
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)

        # Mapping: B=1, F=5, J=9, AE=30, T=19, AU=46, AT=45
        # We grab the necessary columns from Master
        mast_subset = df_m.iloc[:, [1, 5, 9, 30, 19, 46, 45]].copy()
        mast_subset.columns = ['Roll_No', 'Name', 'Sec_J', 'Father', 'Address', 'Phone_AU', 'Phone_AT']
        
        # Clean Master keys for matching
        mast_subset['Match_B'] = mast_subset['Roll_No'].astype(str).str.replace('.0', '', regex=False).str.strip().str.upper()
        mast_subset['Match_C'] = mast_subset['Name'].astype(str).str.strip().str.upper()
        mast_subset['Match_G'] = mast_subset['Sec_J'].astype(str).str.strip().str.upper()

        # 3. Perform Matching (Inner Join)
        df_matched = pd.merge(
            mast_subset, 
            shortage_keys, 
            on=['Match_B', 'Match_C', 'Match_G'], 
            how='inner'
        )

        # 4. Remove Duplicates (Ensure 1 label per student)
        df_matched = df_matched.drop_duplicates(subset=['Roll_No']).copy()

        if st.button("Generate Sorted A4 Label Sheet"):
            if not df_matched.empty:
                # --- SORTING ---
                df_matched['sort_rank'] = df_matched['Roll_No'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No'])

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
                    ws.set_row(r_num, 122) # Height

                    # --- Left Label (Col A) ---
                    d = data_list[i]
                    p1 = clean_val(d.get('Phone_AU'))
                    p2 = clean_val(d.get('Phone_AT'))
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
                        pr1 = clean_val(dr.get('Phone_AU'))
                        pr2 = clean_val(dr.get('Phone_AT'))
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
                
                st.success(f"Generated {len(df_matched)} unique labels.")
                st.download_button(
                    label="📥 Download Unique Label Sheet",
                    data=output.getvalue(),
                    file_name="Student_Unique_Labels.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No matches found between Shortage Report and Master Data using (B, C, G) vs (B, F, J).")
                
    except Exception as e:
        st.error(f"Error processing files: {e}")
else:
    st.info("Please upload both files to continue.")
