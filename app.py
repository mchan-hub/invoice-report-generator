
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.formula.translate import Translator
import re
from io import BytesIO

st.set_page_config(page_title="Invoice Management Report Generator", page_icon="🧾", layout="centered")

st.title("🧾 Invoice Management Report Generator")
st.markdown("Upload your **Xero**, **NetSuite**, and **Report Template** files below to automatically generate the consolidated invoice report.")

st.sidebar.header("Instructions")
st.sidebar.markdown(
    """
    1. Upload the **Xero Overdue CSV** file.
    2. Upload the **NetSuite Overdue CSV** file.
    3. Upload your **Excel Template (.xlsx)**.
    4. Click **Generate Report**.
    5. Download the final `.xlsx` file.
    """
)

xero_file = st.file_uploader("1. Upload Xero CSV", type=["csv"])
ns_file = st.file_uploader("2. Upload NetSuite CSV", type=["csv"])
template_file = st.file_uploader("3. Upload Report Template (.xlsx)", type=["xlsx"])

def clean_amount(val):
    if pd.isna(val): return 0.0
    val_str = str(val)
    cleaned = re.sub(r'[^\d.-]', '', val_str)
    try:
        if cleaned.count('.') > 1:
            parts = cleaned.rsplit('.', 1)
            cleaned = parts[0].replace('.', '') + '.' + parts[1]
        return float(cleaned) if cleaned else 0.0
    except ValueError:
        return 0.0

if st.button("🚀 Generate Report", type="primary"):
    if xero_file and ns_file and template_file:
        try:
            with st.spinner("Processing files..."):
                # 1. Process Xero
                xero_df = pd.read_csv(xero_file)
                xero_mapped = pd.DataFrame({
                    'Location': 'Xero', 
                    'Invoice #': xero_df.get('InvoiceNumber', ''),
                    'Client': xero_df.get('ContactName', ''),
                    'Currency': xero_df.get('Currency', ''),
                    'Amount': xero_df.get('Total', 0),
                    'Due date': pd.to_datetime(xero_df.get('DueDate', ''), format='mixed', dayfirst=True, errors='coerce')
                })

                # 2. Process NetSuite
                ns_df = pd.read_csv(ns_file)
                ns_mapped = pd.DataFrame({
                    'Location': ns_df.get('Subsidiary', 'NetSuite'),
                    'Invoice #': ns_df.get('Document Number', ''),
                    'Client': ns_df.get('Name', ''),
                    'Currency': ns_df.get('Currency', ''),
                    'Amount': ns_df.get('Amount (Foreign Currency)', '').apply(clean_amount),
                    'Due date': pd.to_datetime(ns_df.get('Due Date/Receive By', ''), format='mixed', errors='coerce')
                })

                # Combine Data
                combined_df = pd.concat([xero_mapped, ns_mapped], ignore_index=True)
                combined_df['Due date'] = combined_df['Due date'].dt.strftime('%Y-%m-%d').fillna('')

                # 3. Write to Excel
                wb = openpyxl.load_workbook(template_file)
                ws = wb.active

                header_row = 2
                headers = {ws.cell(row=header_row, column=c).value: c for c in range(1, ws.max_column + 1)}

                col_map = {
                    'Location': headers.get('Location'),
                    'Invoice #': headers.get('Invoice #'),
                    'Client': headers.get('Client'),
                    'Currency': headers.get('Currency'),
                    'Amount': headers.get('Amount'),
                    'Due date': headers.get('Due date'),
                }

                ref_row = 3
                formula_cols = [c for c in range(1, ws.max_column + 1) if ws.cell(row=ref_row, column=c).data_type == 'f']

                start_row = 3
                for r_idx, row in combined_df.iterrows():
                    current_row = start_row + r_idx
                    
                    for col_name, col_idx in col_map.items():
                        if col_idx:
                            ws.cell(row=current_row, column=col_idx).value = row[col_name]
                    
                    if current_row > ref_row:
                        for c_idx in formula_cols:
                            ref_cell = ws.cell(row=ref_row, column=c_idx)
                            new_cell = ws.cell(row=current_row, column=c_idx)
                            new_cell.value = Translator(ref_cell.value, origin=ref_cell.coordinate).translate_formula(new_cell.coordinate)

                # Save to BytesIO object for download
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
            st.success("✅ Report generated successfully!")
            
            st.download_button(
                label="📥 Download Updated Invoice Report",
                data=output,
                file_name="Updated_Invoice_Management_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
            st.error("Please ensure your CSV headers and Excel template format match the requirements.")
    else:
        st.warning("⚠️ Please upload all three files (Xero CSV, NetSuite CSV, and Excel Template) before generating.")
