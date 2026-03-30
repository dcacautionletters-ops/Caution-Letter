import streamlit as st
import pandas as pd
import io

# --- Helper: Clean and Standardize for Matching ---
def clean_match_col(val):
    return str(val).strip().upper().split('.')[0]

# --- Helper: Clean display values (Handles NaN and .0) ---
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan': 
        return ""
    # Removes .0 from numbers and strips whitespace
    return str(val).split('.')[0].strip()

st.set_page_config(page_title="Student Label Generator", layout="wide")
st.title("🏷️ Final Student Label Generator")

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
            # Using openpyxl engine for better Excel compatibility
            df_s = pd.read_excel(file_shortage, skiprows=skip_rows-1)
        
        # Mapping: B=1 (Roll), C=2 (Name), G=6 (Extra/Section)
        df_s_subset = df_s.iloc[:, [1, 2, 6]].copy()
        df_s_subset.columns = ['S_Roll', 'S_Name', 'S_Extra']
        
        # --- UNIQUENESS FIX ---
        # Create a clean key and drop duplicates so 1 student = 1 label
        df_s_subset['S_Roll_key'] = df_s_subset['S_Roll'].apply(clean_match_col)
        df_s_unique = df_s_subset.drop_duplicates(subset=['S_Roll_key']).copy()

        # 2. LOAD MASTER DATA (B=1, F=5, J=9, AE=30, T=19, AU=46, AT=45)
        if file_master.name.endswith('csv'):
            df_m = pd.read_csv(file_master)
        else:
            df_m = pd.read_excel(file_master)
            
        # Mapping to your requested columns
        df_m_subset = df_m.iloc[:, [1, 5, 9, 30, 19, 46, 45]].copy()
        df_m_subset.columns = ['M_Roll', 'M_Name', 'M_Extra', 'To_Name', 'Address', 'Phone1', 'Phone2']
        
        # Create matching keys for Master
        df_m_subset['M_Roll_key'] = df_m_subset['M_Roll'].apply(clean_match_col)
        df_m_subset['M_Name_key'] = df_m_subset['M_Name'].apply(clean_match_col)
        df_m_subset['M_Extra_key'] = df_m_subset['M_Extra'].apply(clean_match_col)

        # 3. MULTI-COLUMN MATCH (B=B, C=F, G=J)
        # Note: We match the Unique Shortage list against the Master
        df_matched = pd.merge(
            df_s_unique, 
            df_m_subset, 
            left_on=['S_Roll_key'], # Primary match on Roll No
            right_on=['M_Roll_key'],
            how='inner'
        )

        st.info(f"Summary: Found {len(df_s_unique)} unique students. Matched {len(df_matched)} with Master Data.")

        if st.button("Generate Final Label Sheet"):
            if not df_matched.empty:
                # Prepare Excel Buffer
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    ws = writer.book.add_worksheet('LABELS')
                    
                    # Page Layout Setup
                    ws.set_column('A:A', 46.5) # Width of label
                    ws.set_column('B:B', 2.5)  # Gap
                    ws.set_column('C:C', 46.5) # Width of label
                    
                    label_fmt = writer.book.add_format({
                        'font_name': 'Calibri', 'font_size': 10, 'text_wrap': True,
                        'valign': 'vcenter', 'border': 1, 'border_color': '#D3D3D3'
                    })

                    data_list = df_matched.to_dict('records')
                    row_cursor = 0
                    
                    # Loop through data 2-at-a-time for side-by-side labels
                    for i in range(0, len(data_list), 2):
                        ws.set_row(row_cursor, 125) # Height of label
                        
                        # --- LEFT LABEL ---
                        d_left = data_list[i]
                        contact_l = f"{clean_val(d_left['Phone1'])} / {clean_val(d_left['Phone2'])}".strip(" / ")
                        txt_left = (
                            f"From: {from_address}\n"
                            f"To, Shri/Smt. {clean_val(d_left['To_Name'])}\n"
                            f"c/o: {clean_val(d_left['S_Name'])}\n"
                            f"Address: {clean_val(d_left['Address'])}\n"
                            f"Contact: {contact_l}    Std ID: {clean_val(d_left['S_Roll'])}"
                        )
                        ws.write(row_cursor, 0, txt_left, label_fmt)

                        # --- RIGHT LABEL ---
                        if i + 1 < len(data_list):
                            d_right = data_list[i+1]
                            contact_r = f"{clean_val(d_right['Phone1'])} / {clean_val(d_right['Phone2'])}".strip(" / ")
                            txt_right = (
                                f"From: {from_address}\n"
                                f"To, Shri/Smt. {clean_val(d_right['To_Name'])}\n"
                                f"c/o: {clean_val(d_right['S_Name'])}\n"
                                f"Address: {clean_val(d_right['Address'])}\n"
                                f"Contact: {contact_r}    Std ID: {clean_val(d_right['S_Roll'])}"
                            )
                            ws.write(row_cursor, 2, txt_right, label_fmt)

                        row_cursor += 1
                        ws.set_row(row_cursor, 8) # Spacer row
                        row_cursor += 1

                    # A4 Print Settings
                    ws.set_paper(9) # A4
                    ws.set_margins(0.3, 0.3, 0.3, 0.3)
                    ws.center_horizontally()

                st.success(f"Done! {len(df_matched)} labels generated.")
                st.download_button(
                    label="📥 Download Labels (A4)",
                    data=output.getvalue(),
                    file_name="Student_Labels_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No matches found. Please check your file columns and start row.")

    except Exception as e:
        st.error(f"Error processing files: {e}")
