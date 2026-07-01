import streamlit as st  # यहाँ सुधारा गया है (as st)
import pandas as pd
import zipfile
import shutil
import os
from datetime import datetime
import io

# Page Configuration (Mobile Friendly Layout)
st.set_page_config(page_title="Billing Merge Tool", page_icon="📊", layout="centered")

st.title("📊 BILLED UNBILLED MERGE")
st.write("Developed by: Anurag Shakya")

# 1. Target Columns
TARGET_COLUMNS = [
    'ACCT_ID', 'METER_SERIAL_NBR', 'METER_READ_REMARK', 
    'MR_SOURCE_CD', 'BILL_BASIS', 'DUE_DATE', 
    'LAST_PAY_DATE', 'AMOUNT_PAYABLE', 'BILL_TYP', 
    'BILL_DATE', 'STATUS'
]

# 2. File Uploader Component (अब यह बॉक्स स्क्रीन पर बिल्कुल सही दिखेगा)
uploaded_files = st.file_uploader(
    "1. Select Files (CSV/Excel/Zip)", 
    type=["xlsx", "xls", "csv", "zip"], 
    accept_multiple_files=True
)

if st.button("GENERATE REPORT (FAST)", type="primary"):
    if not uploaded_files:
        st.warning("कृपया पहले फ़ाइलें सेलेक्ट करें!")
    else:
        with st.spinner("प्रोसेसिंग चल रही है... कृपया इंतज़ार करें..."):
            all_dfs = []
            temp_dir = "temp_streamlit_ver"
            
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)

            try:
                # फ़ाइलों को कलेक्ट करना
                working_files = []
                for uploaded_file in uploaded_files:
                    file_name = uploaded_file.name.upper()
                    
                    # अगर ज़िप फ़ाइल है
                    if uploaded_file.name.lower().endswith('.zip'):
                        zip_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(zip_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        
                        extract_path = os.path.join(temp_dir, "extracted")
                        with zipfile.ZipFile(zip_path, 'r') as z:
                            z.extractall(extract_path)
                            
                        for r, _, files in os.walk(extract_path):
                            for f in files:
                                if f.lower().endswith(('.xlsx', '.csv')):
                                    working_files.append((os.path.join(r, f), f.upper()))
                    else:
                        # नॉर्मल एक्सेल/सीएसवी फ़ाइल
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        working_files.append((file_path, file_name))

                # डेटा प्रोसेस करना (तेज़ लॉजिक)
                for fpath, fname in working_files:
                    st.text(f"पढ़ रहा हूँ: {fname}")
                    
                    if fpath.lower().endswith('.csv'):
                        df = pd.read_csv(fpath, low_memory=False)
                    else:
                        df = pd.read_excel(fpath, engine='openpyxl')

                    # फ़िल्टर्स (Mainpuri Circle)
                    if 'CIRCLE_NAME' in df.columns:
                        df = df[df['CIRCLE_NAME'].astype(str).str.contains('MAINPURI', case=False, na=False)]
                    
                    if df.empty:
                        continue

                    if 'ACCT_ID' in df.columns:
                        df['ACCT_ID'] = df['ACCT_ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(10)

                    if 'METER_SERIAL_NBR' not in df.columns and 'MTR_SRL_NO' in df.columns:
                        df['METER_SERIAL_NBR'] = df['MTR_SRL_NO']
                    
                    if 'SANCTION_LOAD' in df.columns:
                        df['SANCTION_LOAD'] = pd.to_numeric(df['SANCTION_LOAD'], errors='coerce').fillna(0)
                        df = df[df['SANCTION_LOAD'] >= 5]

                    if 'TARIFF_TYPE' in df.columns:
                        df = df[~df['TARIFF_TYPE'].astype(str).str.contains('LMV-5|LMV5', case=False, na=False)]

                    is_unbilled = "UNBILLED" in fname
                    for col in TARGET_COLUMNS:
                        if col not in df.columns: 
                            df[col] = ""
                        if is_unbilled and col in ['LAST_PAY_DATE', 'AMOUNT_PAYABLE', 'BILL_DATE', 'DUE_DATE']:
                            df[col] = ""
                    
                    df_final = df[TARGET_COLUMNS].copy()
                    df_final['STATUS'] = "UNBILLED" if is_unbilled else "BILLED"
                    all_dfs.append(df_final)

                if all_dfs:
                    final_combined = pd.concat(all_dfs, ignore_index=True)
                    
                    # एक्सेल फ़ाइल को मेमोरी में राइट करना
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        final_combined.to_excel(writer, index=False, sheet_name='Data')
                        workbook = writer.book
                        worksheet = writer.sheets['Data']
                        max_row, max_col = final_combined.shape
                        worksheet.add_table(0, 0, max_row, max_col - 1, {
                            'columns': [{'header': col} for col in final_combined.columns],
                            'style': 'Table Style Medium 9'
                        })
                        worksheet.set_column(0, max_col-1, 16)
                    
                    st.success("🎉 मर्ज कम्प्लीट हो गया!")
                    
                    # 3. Download Button
                    out_name = f"Billed_Unbilled_Merze_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.xlsx"
                    st.download_button(
                        label="📥 डाउनलोड एक्सेल रिपोर्ट",
                        data=output.getvalue(),
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("कोई मैचिंग डेटा नहीं मिला (मैनपुरी फ़िल्टर चेक करें)।")

            except Exception as e:
                st.error(f"एरर आई: {str(e)}")
            finally:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)