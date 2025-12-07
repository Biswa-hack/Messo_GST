import streamlit as st
import pandas as pd
import io
import zipfile 
from openpyxl import load_workbook 
from openpyxl.utils.dataframe import dataframe_to_rows

# --- GLOBAL MAPPING ---
COLUMN_MAPPING = {
    'order_date': 'order_date',
    'sub_order_num': 'order_num',
    'hsn_code': 'hsn_code',
    'gst_rate': 'gst_rate',
    'total_taxable_sale_value': 'tcs_taxable_amount',
    'end_customer_state_new': 'end_customer_state_new',
    'quantity': 'QTY',
}


# --- HELPER FUNCTION: PROCESS SINGLE FILE ---
def process_file(file_data, data_type):
    """Processes a single file (sales or returns) from the ZIP archive."""
    
    if file_data is None: return None
    
    # Read Excel file from the provided data stream (from the zip archive)
    df = pd.read_excel(file_data) 
    
    # 1. Apply column renaming using the global map
    df_processed = df.rename(columns=COLUMN_MAPPING)
    required_cols = list(COLUMN_MAPPING.values())
    required_cols_present = [col for col in required_cols if col in df_processed.columns]
    
    if len(required_cols_present) != len(required_cols):
        st.warning(f"⚠️ Input file for {data_type} is missing some required columns.")
    
    df_final = df_processed[required_cols_present].copy()
    df_final.loc[:, 'TYPE'] = data_type
    
    # 2. Handle negative values for Returns 
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

    # 3. Final 8-column order (B to I)
    final_order = ['order_date', 'order_num', 'hsn_code', 'gst_rate', 
                   'tcs_taxable_amount', 'end_customer_state_new', 'TYPE', 'QTY']
    
    final_order_present = [col for col in final_order if col in df_final.columns]
    
    return df_final[final_order_present]


# --- MAIN FUNCTION: PROCESS ZIP AND COMBO ---
def process_zip_and_combine_data(zip_file_uploader, combo_template_file):
    """Handles zip extraction, file identification, processing, and Excel output."""
    
    sales_file_data = None
    returns_file_data = None
    
    # 1. Unzip and Identify Files
    try:
        with zipfile.ZipFile(io.BytesIO(zip_file_uploader.read())) as z:
            for name in z.namelist():
                if name.endswith('.xlsx') or name.endswith('.xls'):
                    if 'return' in name.lower() or 'rtn' in name.lower():
                        returns_file_data = z.open(name)
                    else:
                        sales_file_data = z.open(name)
            
            if not sales_file_data or not returns_file_data:
                 st.error("❌ Could not identify both 'Sales' and 'Returns' files inside the ZIP.")
                 return None

    except zipfile.BadZipFile:
        st.error("❌ The uploaded file is not a valid ZIP archive.")
        return None
    except Exception as e:
        st.error(f"❌ An unexpected error occurred during zip processing: {e}")
        return None

    # 2. Process Data and Merge
    df_sales = process_file(sales_file_data, 'Sale')
    df_returns = process_file(returns_file_data, 'Return')
    
    merged_dfs = []
    if df_sales is not None: merged_dfs.append(df_sales)
    if df_returns is not None: merged_dfs.append(df_returns)

    if not merged_dfs:
        st.error("❌ No valid sales or returns data was processed.")
        return None
        
    df_merged = pd.concat(merged_dfs, ignore_index=True)
    df_merged = df_merged.dropna(axis=1, how='all')

    # 3. Insert Merged Data into Combo Template
    try:
        wb = load_workbook(combo_template_file)
        if 'raw' not in wb.sheetnames:
            st.error("❌ Error: The 'combo' file must contain a sheet named 'raw'.")
            return None
            
        ws = wb['raw']

        # --- A. CLEAR CELLS B3:I[MAX_ROW] (Preserve Formulas in J-O) ---
        start_row_to_clear = 3
        if ws.max_row >= start_row_to_clear:
            rows_to_clear = list(range(start_row_to_clear, ws.max_row + 1))
            for row_idx in rows_to_clear:
                for col_idx in range(2, 10): # Columns 2 (B) through 9 (I)
                    ws.cell(row=row_idx, column=col_idx).value = None
            st.info(f"Cleared old data from columns B:I starting at row 3.")

        # --- B. Paste Merged Data starting at B3 ---
        for r_idx, row in enumerate(dataframe_to_rows(df_merged, header=False, index=False)):
            for c_idx, value in enumerate(row):
                ws.cell(row=start_row_to_clear + r_idx, column=2 + c_idx, value=value)
        
        new_max_row = start_row_to_clear + len(df_merged) - 1
        st.success(f"Successfully pasted {len(df_merged)} rows (B3 to I{new_max_row}).")
        
        # --- C. HARDCODED FORMULAS (K3:O[LastRow]) ---
        formula_target_start_row = 3
        
        # Hardcoded formulas (using {0} as a placeholder for the row index)
        # Note: E1% has been replaced with E{0}/100 for correct openpyxl syntax.
        RAW_FORMULAS = [
            # K: =IF(J1=$X$22,F1*E1%/2,0)
            '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)', 
            # L: =IF(J1=$X$22,F1*E1%/2,0)
            '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)',
            # M: =IF(J1=$X$22,0,F1*E1%)
            '=IF(J{0}=$X$22,0,F{0}*E{0}/100)',
            # N: =K1+L1+M1+F1 
            '=K{0}+L{0}+M{0}+F{0}',
            # O: =(M1+L1+K1)/F1 
            '=(M{0}+L{0}+K{0})/F{0}'
        ]
        
        # Formulas will be pasted starting at Column K (index 11)
        formula_start_col = 11 

        if len(df_merged) > 0:
            for row_idx in range(formula_target_start_row, new_max_row + 1): 
                for col_idx, formula_template in enumerate(RAW_FORMULAS):
                    
                    # 1. Format the formula string with the current row index
                    formatted_formula = formula_template.format(row_idx)

                    # 2. Paste the formula into the cell
                    target_col = formula_start_col + col_idx
                    ws.cell(row=row_idx, column=target_col, value=formatted_formula)
        
            st.info(f"Hardcoded formulas (K:O) applied to {new_max_row - formula_target_start_row + 1} rows.")
            
        st.warning("⚠️ **Pivot Table Refresh:** Please ensure the Pivot Tables are set to 'Refresh data when opening the file' in Excel.")
        
        output = io.BytesIO()
        wb.save(output)
        
        return output.getvalue()
        
    except Exception as e:
        st.error(f"❌ An error occurred during file manipulation: {e}")
        return None

# ==============================================================================
# Streamlit UI
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
        type=['zip'],
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

st.sidebar.markdown("## 📚 Guidance")
st.sidebar.markdown("---")
st.sidebar.warning("**Reminder:** The Pivot Tables will **not** refresh until you open the file in Excel and confirm the refresh due to cloud environment limitations.")
