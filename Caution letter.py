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
st.title("🏷️ Student Label Generator (Unique 50 Students Fix)")

# --- FILE UPLOAD ---
col1, col2 = st.columns(2)
with col1:
    file_shortage = st.file_uploader("Upload Shortage File", type=['xlsx', 'csv'])
    skip_rows = st.number_input("Data starts on row:", min_value=1, value=4)
with col2:
    file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_shortage and file_master:
    try:
        # 1. LOAD SHORTAGE REPORT (B=1, C=2, G=6)
        df_s = pd.read_excel(file_shortage, skiprows=skip_rows-1) if file_shortage.name.endswith('xlsx') else pd.read_csv(file_shortage, skiprows=skip_rows-1)
        
        # Select columns and Rename
        df_s_clean = df_s.iloc[:, [1, 2, 6]].copy()
        df_s_clean.columns = ['S_Roll', 'S_Name', 'S_Extra']

        # --- THE FIX: DROP DUPLICATES ---
        # This ensures that even if a student is listed 10 times in the shortage report, 
        # they only get ONE label.
        df_s_clean['S_Roll_key'] = df_s_clean['S_Roll'].apply(clean_match_col)
        df_s_unique = df_s_clean.drop_duplicates(subset=['S_Roll_key']).copy()

        # 2. LOAD MASTER DATA (B=1, F=5, J=9, AE=30, T=19, AU=46, AT=45)
        df_m = pd.read_excel(file_master) if file_master.name.endswith('xlsx') else pd.read_csv(file_master)
        df_m_clean = df_m.iloc[:, [1, 5, 9, 30, 19, 46, 45]].copy()
        df_m_clean.columns = ['M_Roll', 'M_Name', 'M_Extra', 'To_Name', 'Address', 'Phone1', 'Phone2']
        df_m_clean['M_Roll_key'] = df_m_clean['M_Roll'].apply(clean_match_col)

        # 3. MATCH (Using only the unique list)
        df_matched = pd.merge(
            df_s_unique, 
            df_m_clean, 
            left_on='S_Roll_key',
            right_on='M_Roll_key',
            how='inner'
        )

        st.write(f"### Found {len(df_s_unique)} unique students in Shortage Report.")
        st.write(f"### Successfully matched {len(df_matched)} students with Master Data.")

        if st.button("Generate Labels"):
            # ... (Rest of the Excel generation code remains the same)
            # Use 'df_matched' to generate the 50 labels.
            pass

    except Exception as e:
        st.error(f"Error: {e}")
