import streamlit as st
import pandas as pd
import io

# --- Helper Function: Clean Phone Numbers & IDs ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    # Remove .0 and ensure it's a clean string
    text = str(val).split('.')[0].strip()
    return text

# --- Custom Sorting Function ---
def get_sort_rank(roll):
    roll = str(roll).upper().strip()
    if roll.startswith('25CG'): return 1
    if roll.startswith('25CAI'): return 2
    if roll.startswith('25CDS'): return 3
    if roll.startswith('24C'): return 4
    if roll.startswith('23C'): return 5
    return 6 

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Robust Version)")

# --- SIDEBAR SETTINGS ---
st.sidebar.header("Global Settings")
from_address = st.sidebar.text_area(
    "Edit 'From' Address:", 
    value="Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Attendance/Shortage Report")
    file_caution = st.file_uploader("Upload Attendance File", type=['xlsx', 'csv'], key="caution")
    # Increase this if your Excel has many titles/logos at the top
    skip_rows = st.number_input("Data (Headers) starts on row:", min_value=1, value=4)

with col2:
    st.subheader("2. Master Database")
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'], key="master")

if file_caution and file_master:
    try:
        # --- LOADING ATTENDANCE ---
        if file_caution.name.endswith('csv'):
            df_c = pd.read_csv(file_caution, skiprows=skip_rows-1)
        else:
            df_c = pd.read_excel(file_caution, skiprows=skip_rows-1)
        
        # Clean Attendance Roll Numbers (Column Index 1 = Col B)
        # We force to Uppercase and Strip spaces to prevent "invisible" mismatches
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.split('.').str[0].str.strip().str.upper().unique()
        
        # --- LOADING MASTER ---
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)

        # Map Master Data (B=1, F=5, AD=29, S=18, AT=45, AS=44)
        # Using a dictionary to catch errors if columns are missing
        indices = [1, 5, 29, 18, 45, 44]
        mast_subset = df_m.iloc[:, indices].copy()
        mast_subset.columns = ['Roll_No', 'Name', 'Father', 'Address', 'Father_Phone', 'Student_Phone']
        
        # Clean Master Roll Numbers for matching
        mast_subset['Roll_No_Clean'] = mast_subset['Roll_No'].astype(str).str.split('.').str[0].str.strip().str.upper()
        
        # --- MATCHING ---
        df_matched = mast_subset[mast_subset['Roll_No_Clean'].isin(caution_rolls)].copy()

        # Debugging Info (Hidden by default, good for troubleshooting)
        with st.expander("🔍 View Data Match Stats"):
            st.write(f"Unique IDs in Attendance: {len(caution_rolls)}")
            st.write(f"Matches found in Master: {len(df_matched)}")
            if len(df_matched) == 0:
                st.warning("No matches found! Check if 'Attendance row start' is correct.")

        if st.button("Generate Sorted A4 Label Sheet"):
            if not df_matched.empty:
                # Sort by series (25CG, etc.)
                df_matched['sort_rank'] = df_matched['Roll_No_Clean'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'Roll_No_Clean'])

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as workbook_writer:
                    ws = workbook_writer.book.add_worksheet('PRINT_LABELS')

                    # Formatting
                    ws.set_column('A:A', 46.5)
                    ws.set_column('B:B', 2.5) 
                    ws.set_column('C:C', 46.5)
                    
                    label_fmt = workbook_writer.book.add_format({
                        'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True,
                        'valign': 'vcenter', 'border': 1, 'border_color': '#C8C8C8'
                    })

                    r_num = 0 
                    data_list = df_matched.to_dict('records')
                    
                    for i in range(0, len(data_list), 2):
                        ws.set_row(r_num, 122)
                        
                        # Left Label
                        d = data_list[i]
                        contact = f"{clean_val(d['Father_Phone'])} / {clean_val(d['Student_Phone'])}".strip(" / ")
                        txt_left = (f"From: {from_address}\n"
                                    f"To, Shri/Smt. {clean_val(d['Father'])}\n"
                                    f"c/o: {clean_val(d['Name'])}\n"
                                    f"Address: {clean_val(d['Address'])}\n"
                                    f"Contact: {contact}   ID: {d['Roll_No_Clean']}")
                        ws.write(r_num, 0, txt_left, label_fmt)

                        # Right Label
                        if i + 1 < len(data_list):
                            dr = data_list[i+1]
                            contact_r = f"{clean_val(dr['Father_Phone'])} / {clean_val(dr['Student_Phone'])}".strip(" / ")
                            txt_right = (f"From: {from_address}\n"
                                         f"To, Shri/Smt. {clean_val(dr['Father'])}\n"
                                         f"c/o: {clean_val(dr['Name'])}\n"
                                         f"Address: {clean_val(dr['Address'])}\n"
                                         f"Contact: {contact_r}   ID: {dr['Roll_No_Clean']}")
                            ws.write(r_num, 2, txt_right, label_fmt)

                        r_num += 1
                        ws.set_row(r_num, 8) # Gap
                        r_num += 1

                    ws.set_paper(9) # A4
                    ws.set_margins(0.2, 0.2, 0.3, 0.3)
                    ws.center_horizontally()

                st.success(f"Success! Created {len(df_matched)} labels.")
                st.download_button("📥 Download Labels", output.getvalue(), "Student_Labels.xlsx")
            else:
                st.error("No students matched. Verify your Attendance file row starting point.")

    except Exception as e:
        st.error(f"Error processing files: {e}")
