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

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Student Label Generator (Unique Labels Only)")

# --- UPLOAD SECTION ---
col1, col2 = st.columns(2)
with col1:
    file_shortage = st.file_uploader("1. Upload Shortage Report", type=['xlsx', 'csv'])
    skip_rows = st.number_input("Shortage data starts on row:", min_value=1, value=4)
with col2:
    file_master = st.file_uploader("2. Upload Master Data Set", type=['xlsx', 'csv'])

if file_shortage and file_master:
    try:
        # 1. LOAD SHORTAGE REPORT (B=1, C=2, G=6)
        df_s = pd.read_excel(file_shortage, skiprows=skip_rows-1) if file_shortage.name.endswith('xlsx') else pd.read_csv(file_shortage, skiprows=skip_rows-1)
        
        # Keep B, C, G
        df_s_subset = df_s.iloc[:, [1, 2, 6]].copy()
        df_s_subset.columns = ['S_Roll', 'S_Name', 'S_Extra']
        
        # --- CRITICAL FIX: FORCE UNIQUENESS ---
        # This line removes all duplicates. 209 rows will become 50 rows here.
        df_s_subset['S_Roll_key'] = df_s_subset['S_Roll'].apply(clean_match_col)
        df_s_unique = df_s_subset.drop_duplicates(subset=['S_Roll_key']).copy()

        # 2. LOAD MASTER DATA (B=1, F=5, J=9, AE=30, T=19, AU=46, AT=45)
        df_m = pd.read_excel(file_master) if file_master.name.endswith('xlsx') else pd.read_csv(file_master)
        df_m_subset = df_m.iloc[:, [1, 5, 9, 30, 19, 46, 45]].copy()
        df_m_subset.columns = ['M_Roll', 'M_Name', 'M_Extra', 'To_Name', 'Address', 'Phone1', 'Phone2']
        df_m_subset['M_Roll_key'] = df_m_subset['M_Roll'].apply(clean_match_col)

        # 3. MATCH UNIQUE STUDENTS ONLY
        df_matched = pd.merge(
            df_s_unique, 
            df_m_subset, 
            left_on='S_Roll_key', 
            right_on='M_Roll_key', 
            how='inner'
        )

        # Dashboard to show you it's working
        st.success(f"Step 1: Found {len(df_s_subset)} total entries.")
        st.success(f"Step 2: Filtered down to {len(df_s_unique)} unique students.")
        st.success(f"Step 3: Matched {len(df_matched)} students with Master Address Data.")

        if st.button("Generate Final 50 Labels"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                ws = writer.book.add_worksheet('LABELS')
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.5)
                ws.set_column('C:C', 46.5)
                label_fmt = writer.book.add_format({'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True, 'valign': 'vcenter', 'border': 1})

                data = df_matched.to_dict('records')
                row_idx = 0
                for i in range(0, len(data), 2):
                    ws.set_row(row_idx, 125)
                    # Left Label
                    d = data[i]
                    txt = f"To: Shri/Smt. {clean_val(d['To_Name'])}\nc/o: {clean_val(d['S_Name'])}\nAddress: {clean_val(d['Address'])}\nContact: {clean_val(d['Phone1'])} / {clean_val(d['Phone2'])}\nID: {clean_val(d['S_Roll'])}"
                    ws.write(row_idx, 0, txt, label_fmt)
                    
                    # Right Label
                    if i + 1 < len(data):
                        dr = data[i+1]
                        txt_r = f"To: Shri/Smt. {clean_val(dr['To_Name'])}\nc/o: {clean_val(dr['S_Name'])}\nAddress: {clean_val(dr['Address'])}\nContact: {clean_val(dr['Phone1'])} / {clean_val(dr['Phone2'])}\nID: {clean_val(dr['S_Roll'])}"
                        ws.write(row_idx, 2, txt_r, label_fmt)
                    
                    row_idx += 2 # Move down for next pair

            st.download_button("📥 Download Unique Labels", output.getvalue(), "Unique_Student_Labels.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
