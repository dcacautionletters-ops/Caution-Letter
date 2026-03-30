import streamlit as st
import pandas as pd
import io

# --- Helper: Clean and Standardize for Matching ---
def clean_match_col(val):
    return str(val).strip().upper().split('.')[0]

# --- Helper: Clean display values ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    return str(val).split('.')[0].strip()

# --- Custom Sorting ---
def get_sort_rank(roll):
    roll = str(roll).upper()
    if '25CG' in roll: return 1
    if '25CAI' in roll: return 2
    if '25CDS' in roll: return 3
    if '24C' in roll: return 4
    if '23C' in roll: return 5
    return 6 

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Precise Match Version)")

# --- SIDEBAR ---
st.sidebar.header("Global Settings")
from_address = st.sidebar.text_area(
    "Edit 'From' Address:", 
    value="Presidency College Bangalore (AUTONOMOUS)\nKempapura, Hebbal, Bengaluru - 560024"
)

# --- UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Shortage Report")
    file_shortage = st.file_uploader("Upload Shortage File", type=['xlsx', 'csv'])
    skip_rows = st.number_input("Data starts on row (Header row):", min_value=1, value=4)

with col2:
    st.subheader("2. Master Data Set")
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_shortage and file_master:
    try:
        # 1. LOAD SHORTAGE REPORT (B=1, C=2, G=6)
        if file_shortage.name.endswith('csv'):
            df_s = pd.read_csv(file_shortage, skiprows=skip_rows-1)
        else:
            df_s = pd.read_excel(file_shortage, skiprows=skip_rows-1)
        
        # Keep only required columns and rename for clarity
        df_s_clean = df_s.iloc[:, [1, 2, 6]].copy()
        df_s_clean.columns = ['S_Roll', 'S_Name', 'S_Extra']
        
        # 2. LOAD MASTER DATA (B=1, F=5, J=9, AE=30, T=19, AU=46, AT=45)
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)
            
        df_m_clean = df_m.iloc[:, [1, 5, 9, 30, 19, 46, 45]].copy()
        df_m_clean.columns = ['M_Roll', 'M_Name', 'M_Extra', 'To_Name', 'Address', 'Phone1', 'Phone2']

        # 3. PREPARE FOR MERGE (Standardize strings)
        for col in ['S_Roll', 'S_Name', 'S_Extra']:
            df_s_clean[col + '_key'] = df_s_clean[col].apply(clean_match_col)
            
        for col in ['M_Roll', 'M_Name', 'M_Extra']:
            df_m_clean[col + '_key'] = df_m_clean[col].apply(clean_match_col)

        # 4. PERFORM MULTI-COLUMN MATCH
        # Matches B to B, C to F, and G to J
        df_matched = pd.merge(
            df_s_clean, 
            df_m_clean, 
            left_on=['S_Roll_key', 'S_Name_key', 'S_Extra_key'],
            right_on=['M_Roll_key', 'M_Name_key', 'M_Extra_key'],
            how='inner'
        )

        st.write(f"### Match Result: {len(df_matched)} students found.")

        if st.button("Generate Final Label Sheet"):
            if not df_matched.empty:
                # Sorting
                df_matched['sort_rank'] = df_matched['S_Roll_key'].apply(get_sort_rank)
                df_matched = df_matched.sort_values(by=['sort_rank', 'S_Roll_key'])

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    ws = writer.book.add_worksheet('LABELS')
                    
                    # Layout setup
                    ws.set_column('A:A', 46.5)
                    ws.set_column('B:B', 2.5) 
                    ws.set_column('C:C', 46.5)
                    label_fmt = writer.book.add_format({
                        'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True,
                        'valign': 'vcenter', 'border': 1, 'border_color': '#D3D3D3'
                    })

                    data = df_matched.to_dict('records')
                    r_idx = 0
                    
                    for i in range(0, len(data), 2):
                        ws.set_row(r_idx, 125) # Label height
                        
                        # --- LEFT LABEL ---
                        d = data[i]
                        contact = f"{clean_val(d['Phone1'])} / {clean_val(d['Phone2'])}".strip(" / ")
                        txt_left = (
                            f"From: {from_address}\n"
                            f"To, Shri/Smt. {clean_val(d['To_Name'])}\n"
                            f"c/o: {clean_val(d['S_Name'])}\n"
                            f"Address: {clean_val(d['Address'])}\n"
                            f"Contact: {contact}    Std ID: {clean_val(d['S_Roll'])}"
                        )
                        ws.write(r_idx, 0, txt_left, label_fmt)

                        # --- RIGHT LABEL ---
                        if i + 1 < len(data):
                            dr = data[i+1]
                            contact_r = f"{clean_val(dr['Phone1'])} / {clean_val(dr['Phone2'])}".strip(" / ")
                            txt_right = (
                                f"From: {from_address}\n"
                                f"To, Shri/Smt. {clean_val(dr['To_Name'])}\n"
                                f"c/o: {clean_val(dr['S_Name'])}\n"
                                f"Address: {clean_val(dr['Address'])}\n"
                                f"Contact: {contact_r}    Std ID: {clean_val(dr['S_Roll'])}"
                            )
                            ws.write(r_idx, 2, txt_right, label_fmt)

                        r_idx += 1
                        ws.set_row(r_idx, 8) # Vertical Gap
                        r_idx += 1

                    ws.set_paper(9) # A4
                    ws.set_margins(0.3, 0.3, 0.3, 0.3)
                
                st.success("Labels Ready!")
                st.download_button("📥 Download Labels", output.getvalue(), "Final_Student_Labels.xlsx")
            else:
                st.error("No matches found. Check your 'Data starts on row' setting.")

    except Exception as e:
        st.error(f"Error: {e}")
