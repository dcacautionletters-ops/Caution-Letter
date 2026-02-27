import streamlit as st
import pandas as pd
import io

# Function to clean text for Excel
def clean_val(val):
    if pd.isna(val): return ""
    return str(val).strip()

st.title("🏷️ Student Label Generator (Excel Output)")

file_caution = st.file_uploader("Upload Caution Letter File", type=['xlsx', 'csv'])
file_master = st.file_uploader("Upload Master Database", type=['xlsx', 'csv'])

if file_caution and file_master:
    try:
        # Load Files
        df_c = pd.read_csv(file_caution) if file_caution.name.endswith('csv') else pd.read_excel(file_caution)
        df_m = pd.read_csv(file_master) if file_master.name.endswith('csv') else pd.read_excel(file_master)

        # 1. Match Roll No (Caution Col B -> Master 'ROLL NUMBER')
        caution_rolls = df_c.iloc[:, 1].dropna().astype(str).str.strip().unique()
        df_m['ROLL NUMBER'] = df_m['ROLL NUMBER'].astype(str).str.strip()
        df_matched = df_m[df_m['ROLL NUMBER'].isin(caution_rolls)]

        if st.button("Generate Excel Label Sheet"):
            if not df_matched.empty:
                # Create an In-Memory Excel File
                output = io.BytesIO()
                workbook = pd.ExcelWriter(output, engine='xlsxwriter')
                ws = workbook.book.add_worksheet('PRINT_LABELS')

                # --- 2. FORMATTING SETUP (Exactly like VBA) ---
                # Column Widths (VBA 46.5 and 2.5)
                ws.set_column('A:A', 46.5)
                ws.set_column('B:B', 2.5)
                ws.set_column('C:C', 46.5)

                # Cell Format (Calibri, Size 9, Wrapped, V-Center, Border)
                label_format = workbook.book.add_format({
                    'font_name': 'Calibri',
                    'font_size': 10,
                    'text_wrap': True,
                    'valign': 'vcenter',
                    'border': 1,
                    'border_color': '#C8C8C8' # RGB(200, 200, 200)
                })

                # --- 3. LOOP THROUGH DATA (2x6 Grid) ---
                r_num = 0 # Excel Row Index starts at 0
                data_list = df_matched.to_dict('records')
                num_records = len(data_list)
                
                for i in range(0, num_records, 2):
                    # Set Row Height (VBA 122)
                    ws.set_row(r_num, 122)

                    # --- Left Label (Col A) ---
                    row_data = data_list[i]
                    contact = f"{clean_val(row_data.get('STUDENT PHONE'))}/{clean_val(row_data.get('PARENT PHONE'))}".strip("/")
                    
                    # Construct text (No Bold in specific characters via basic write, handled below)
                    txt_left = (f"From: Presidency College Bangalore (AUTONOMOUS)\n"
                                f"To, Shri/Smt. {clean_val(row_data.get('FATHER'))}\n"
                                f"c/o: {clean_val(row_data.get('NAME'))}\n"
                                f"Address: {clean_val(row_data.get('ADDRESS'))}\n"
                                f"Contact: {contact}  ID: {clean_val(row_data.get('ROLL NUMBER'))}")
                    
                    ws.write(r_num, 0, txt_left, label_format)

                    # --- Right Label (Col C) ---
                    if i + 1 < num_records:
                        row_data_r = data_list[i+1]
                        contact_r = f"{clean_val(row_data_r.get('STUDENT PHONE'))}/{clean_val(row_data_r.get('PARENT PHONE'))}".strip("/")
                        
                        txt_right = (f"From: Presidency College Bangalore (AUTONOMOUS)\n"
                                     f"To, Shri/Smt. {clean_val(row_data_r.get('FATHER'))}\n"
                                     f"c/o: {clean_val(row_data_r.get('NAME'))}\n"
                                     f"Address: {clean_val(row_data_r.get('ADDRESS'))}\n"
                                     f"Contact: {contact_r}  ID: {clean_val(row_data_r.get('ROLL NUMBER'))}")
                        
                        ws.write(r_num, 2, txt_right, label_format)

                    # --- Spacing Row (VBA 8 points) ---
                    r_num += 1
                    ws.set_row(r_num, 8)
                    r_num += 1

                # --- 4. PAGE SETUP ---
                ws.set_margins(left=0.2, right=0.2, top=0.2, bottom=0.2) # Approx 0.5cm
                ws.set_paper(9) # A4
                ws.center_horizontally()

                workbook.close()
                st.download_button(
                    label="📥 Download Excel Labels",
                    data=output.getvalue(),
                    file_name="Student_Labels.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success(f"Excel Sheet ready for {num_records} students.")
            else:
                st.error("No matches found.")
    except Exception as e:
        st.error(f"Error: {e}")

