import streamlit as st
import pandas as pd
import io
import zipfile # Needed to handle the compressed file
from openpyxl import load_workbook 
from openpyxl.utils.dataframe import dataframe_to_rows

# (Helper function process_file remains the same as before)
def process_file(file_data, data_type):
    COLUMN_MAPPING = {
        'order_date': 'order_date',
        'sub_order_num': 'order_num',
        'hsn_code': 'hsn_code',
        'gst_rate': 'gst_rate',
        'total_taxable_sale_value': 'tcs_taxable_amount',
        'end_customer_state_new': 'end_customer_state_new',
        'quantity': 'QTY',
    }

    if file_data is None: return None
    
    # Read Excel file from the provided data stream (from the zip archive)
    df = pd.read_excel(file_data) 
    
    df_processed = df.rename(columns=COLUMN_MAPPING)
    required_cols = list(COLUMN_MAPPING.values())
    required_cols_present = [col for col in required_cols if col in df_processed.columns]
    
    if len(required_cols_present) != len(required_cols):
        st.warning(f"⚠️ Input file for {data_type} is missing some required columns.")
    
    df_final = df_processed[required_cols_present].copy()
    
    df_final.loc[:, 'TYPE'] = data_type
    
    # Handle negative values for Returns (Taxable amount AND Quantity)
    if data_type == 'Return':
        
        if 'tcs_taxable_amount' in df_final.columns:
             df_final.loc[:, 'tcs_taxable_amount'] = pd.to_numeric(
                 df_final['tcs_taxable_amount'], errors='coerce'
             ).abs() * -1
        
        if 'QTY' in df_final.columns:
             df_final.loc[:, 'QTY'] = pd.to_numeric(
                 df_final['QTY'], errors='coerce'
             ).abs() * -1
        
    else:
        # Ensure sales values are positive
        if 'tcs_taxable_amount' in df_final.columns:
             df_final.loc[:, 'tcs_taxable_amount'] = pd.to_numeric(
                 df_final['tcs_taxable_amount'], errors='coerce'
             ).abs()
        if 'QTY' in df_final.columns:
             df_final.loc[:, 'QTY'] = pd.to_numeric(
                 df_final['QTY'], errors='coerce'
             ).abs()

    # 🟢 NEW 8-COLUMN FINAL ORDER (B to I) 🟢
    final_order = ['order_date', 'order_num', 'hsn_code', 'gst_rate', 
                   'tcs_taxable_amount', 'end_customer_state_new', 'TYPE', 'QTY']
    
    final_order_present = [col for col in final_order if col in df_final.columns]
    
    return df_final[final_order_present]


def process_zip_and_combine_data(zip_file_uploader, combo_template_file):
    """Handles zip extraction, file identification, processing, and merging."""
    
    sales_file_data = None
    returns_file_data = None
    
    # 1. Unzip and Identify Files
    try:
        # Open the zip file in memory using io.BytesIO
        with zipfile.ZipFile(io.BytesIO(zip_file_uploader.read())) as z:
            
            # Simple heuristic to identify sales vs. returns file based on name
            for name in z.namelist():
                # We only care about Excel files
                if name.endswith('.xlsx') or name.endswith('.xls'):
                    if 'return' in name.lower() or 'rtn' in name.lower():
                        returns_file_data = z.open(name)
                        st.info(f"Identified Returns file: {name}")
                    else:
                        # Assume all other relevant Excel files are sales
                        sales_file_data = z.open(name)
                        st.info(f"Identified Sales file: {name}")
            
            if not sales_file_data or not returns_file_data:
                 st.error("❌ Could not identify both 'Sales' and 'Returns' files inside the ZIP.")
                 return None

    except zipfile.BadZipFile:
        st.error("❌ The uploaded file is not a valid ZIP archive.")
        return None
    except Exception as e:
        st.error(f"❌ An unexpected error occurred during zip processing: {e}")
        return None

    # 2. Process Sales and Returns Data
    df_sales = process_file(sales_file_data, 'Sale')
    df_returns = process_file(returns_file_data, 'Return')
    
    merged_dfs = []
    if df_sales is not None: merged_dfs.append(df_sales)
    if df_returns is not None: merged_dfs.append(df_returns)

    if not merged_dfs:
        st.error("❌ No valid sales or returns data was processed.")
        return None
        
    df_merged = pd.concat(merged_dfs, ignore_index=True)
    
    # Remove columns that are entirely blank (NaN)
    df_merged = df_merged.dropna(axis=1, how='all')

    # 3. Insert Merged Data into Combo Template using openpyxl (Same logic as before)
    try:
        wb = load_workbook(combo_template_file)
        
        if 'raw' not in wb.sheetnames:
            st.error("❌ Error: The 'combo' file must contain a sheet named 'raw'.")
            return None
            
        ws = wb['raw']

        # --- A. Delete existing data below row B3 ---
        start_row_to_clear = 3
        max_row = ws.max_row
        
        if max_row >= start_row_to_clear:
            rows_to_delete = max_row - start_row_to_clear + 1
            ws.delete_rows(start_row_to_clear, rows_to_delete)


        # --- B. Paste Merged Data starting at B3 (Row 3, Column 2) ---
        # Data has 8 columns, fitting B to I.
        for r_idx, row in enumerate(dataframe_to_rows(df_merged, header=False, index=False)):
            for c_idx, value in enumerate(row):
                # Column 2 is B. c_idx runs from 0 to 7 (8 total columns)
                ws.cell(row=start_row_to_clear + r_idx, column=2 + c_idx, value=value)
        
        st.success(f"Successfully pasted {len(df_merged)} rows starting at B3, ending at Column I.")

        st.warning("⚠️ **Pivot Table Refresh:** Refreshing Pivot Tables is NOT possible on this cloud service. Please ensure the Pivot Tables in your template are set to **'Refresh data when opening the file'** in Excel before you download the output.")
        
        output = io.BytesIO()
        wb.save(output)
        
        return output.getvalue()
        
    except Exception as e:
        st.error(f"❌ An error occurred during file manipulation: {e}")
        return None

# ==============================================================================
# Streamlit UI (Modified for ZIP upload)
# ==============================================================================
st.set_page_config(
    page_title="TCS Data Processor",
    layout="wide",
    initial_sidebar_state="auto"
)

st.title("📊 TCS Data Integration & Template Filler")
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Zipped Sales & Returns Data")
    zipped_files = st.file_uploader(
        "Upload a single ZIP file containing both the Sales and Returns Excel sheets",
        type=['zip'], # Only accept zip files
        key='zipped_files'
    )

with col2:
    st.subheader("2. Combo Template")
    combo_template_file = st.file_uploader(
        "Upload the Combo Template (with 'raw' sheet)",
        type=['xlsx', 'xls'],
        key='combo'
    )
    st.info("Template must contain a sheet named **'raw'**.")


st.markdown("---")

# 3. Processing and Download
if zipped_files and combo_template_file:
    st.subheader("3. Process and Download")
    
    if st.button("🚀 Generate Final Combo Report"):
        with st.spinner('Processing ZIP, integrating data, and saving...'):
            processed_excel_data = process_zip_and_combine_data(zipped_files, combo_template_file)

        if processed_excel_data:
            st.download_button(
                label="⬇️ Download Modified Combo Report.xlsx",
                data=processed_excel_data,
                file_name="Modified_Combo_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.balloons()
        else:
            st.error("❌ Failed to process data. Please check file contents and try again.")
