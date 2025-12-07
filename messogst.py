import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook 
from openpyxl.utils.dataframe import dataframe_to_rows

def process_and_combine_data(sales_file, returns_file, combo_template_file):
    """
    Processes the sales/returns, combines them, and inserts the 7 final columns 
    into the 'raw' sheet of the combo template starting at B3 (B through H).
    """
    
    COLUMN_MAPPING = {
        'order_date': 'order_date',
        'sub_order_num': 'order_num',
        'hsn_code': 'hsn_code',
        'gst_rate': 'gst_rate',
        'total_taxable_sale_value': 'tcs_taxable_amount',
        'end_customer_state_new': 'end_customer_state_new',
        'quantity': 'QTY',
    }

    # --- Helper function for processing a single file ---
    def process_file(uploaded_file, data_type):
        if uploaded_file is None: return None
        df = pd.read_excel(uploaded_file)
        
        df_processed = df.rename(columns=COLUMN_MAPPING)
        required_cols = list(COLUMN_MAPPING.values())
        required_cols_present = [col for col in required_cols if col in df_processed.columns]
        
        if len(required_cols_present) != len(required_cols):
            st.warning(f"⚠️ Input file '{uploaded_file.name}' is missing some required columns.")
        
        df_final = df_processed[required_cols_present].copy()
        
        df_final.loc[:, 'TYPE'] = data_type
        
        # Handle negative values for Returns (Taxable amount AND Quantity)
        if data_type == 'Return':
            st.info(f"Applying negative signs to 'tcs_taxable_amount' and 'QTY' for **{data_type}** data.")
            
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

        # Final 7 columns to fit B-H range (order_date is excluded)
        final_order = ['order_num', 'hsn_code', 'gst_rate', 
                       'tcs_taxable_amount', 'end_customer_state_new', 'TYPE', 'QTY']
        
        final_order_present = [col for col in final_order if col in df_final.columns]
        
        return df_final[final_order_present]


    # 3. Processing and Merging
    df_sales = process_file(sales_file, 'Sale')
    df_returns = process_file(returns_file, 'Return')
    
    merged_dfs = []
    if df_sales is not None: merged_dfs.append(df_sales)
    if df_returns is not None: merged_dfs.append(df_returns)

    if not merged_dfs:
        st.error("❌ No valid sales or returns data was processed.")
        return None
        
    df_merged = pd.concat(merged_dfs, ignore_index=True)
    
    # Remove columns that are entirely blank (NaN)
    st.info("Cleaning up merged data: removing columns that are completely empty.")
    df_merged = df_merged.dropna(axis=1, how='all')

    # 4. Insert Merged Data into Combo Template using openpyxl
    try:
        wb = load_workbook(combo_template_file)
        
        if 'raw' not in wb.sheetnames:
            st.error("❌ Error: The 'combo' file must contain a sheet named 'raw'.")
            return None
            
        ws = wb['raw']
        st.info("Template 'raw' sheet loaded.")

        # --- A. Delete existing data below row B3 ---
        start_row_to_clear = 3
        max_row = ws.max_row
        
        if max_row >= start_row_to_clear:
            rows_to_delete = max_row - start_row_to_clear + 1
            ws.delete_rows(start_row_to_clear, rows_to_delete)
            st.info(f"Cleared {rows_to_delete} potential old rows in the 'raw' sheet.")


        # --- B. Paste Merged Data starting at B3 (Row 3, Column 2) ---
        for r_idx, row in enumerate(dataframe_to_rows(df_merged, header=False, index=False)):
            for c_idx, value in enumerate(row):
                ws.cell(row=start_row_to_clear + r_idx, column=2 + c_idx, value=value)
        
        st.success(f"Successfully pasted {len(df_merged)} rows starting at B3, ending at Column H.")

        # --- C. Refresh Pivot Tables (Warning remains) ---
        st.warning("⚠️ **Pivot Table Refresh:** Refreshing Pivot Tables is NOT possible on this cloud service. Please ensure the Pivot Tables in your template are set to **'Refresh data when opening the file'** in Excel before you download the output.")
        
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

col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. Sales Data")
    sales_file = st.file_uploader(
        "Upload the TCS Sales Excel File",
        type=['xlsx', 'xls'],
        key='sales'
    )

with col2:
    st.subheader("2. Returns Data")
    returns_file = st.file_uploader(
        "Upload the TCS Sales Return Excel File",
        type=['xlsx', 'xls'],
        key='returns'
    )

with col3:
    st.subheader("3. Combo Template")
    combo_template_file = st.file_uploader(
        "Upload the Combo Template (with 'raw' sheet)",
        type=['xlsx', 'xls'],
        key='combo'
    )
    st.info("Template must contain a sheet named **'raw'**.")


st.markdown("---")

if sales_file and returns_file and combo_template_file:
    st.subheader("4. Process and Download")
    
    if st.button("🚀 Generate Final Combo Report"):
        with st.spinner('Processing data, integrating into template, and saving...'):
            processed_excel_data = process_and_combine_data(sales_file, returns_file, combo_template_file)

        if processed_excel_data:
            st.download_button(
                label="⬇️ Download Modified Combo Report.xlsx",
                data=processed_excel_data,
                file_name="Modified_Combo_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.balloons()
        else:
            st.error("❌ Failed to process data. Please check logs and file integrity.")

st.sidebar.markdown("## 📚 Guidance")
st.sidebar.markdown("---")
st.sidebar.warning("**Reminder:** The Pivot Tables will **not** refresh until you open the file in Excel and confirm the refresh due to cloud environment limitations.")
