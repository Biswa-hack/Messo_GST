import streamlit as st
import pandas as pd
import io

def process_and_combine_data(sales_file, returns_file):
    """
    Processes the uploaded sales and returns Excel files, applies the 
    specified business rules, and combines them into a multi-sheet Excel file.
    """
    
    # 1. Define the Column Mapping (Input Column Name -> Output Column Name)
    # NOTE: The original request had a conflicting mapping (sub_order_num -> order_date AND sub_order_num -> order_num).
    # This script makes an assumption that both 'order_date' and 'sub_order_num' exist in the input files 
    # to create the 'order_date' and 'order_num' output columns respectively, as this is a more logical structure.
    COLUMN_MAPPING = {
        'order_date': 'order_date',
        'sub_order_num': 'order_num',
        'hsn_code': 'hsn_code',
        'gst_rate': 'gst_rate',
        'total_taxable_sale_value': 'tcs_taxable_amount',
        'end_customer_state_new': 'end_customer_state_new',
        'quantity': 'QTY',
        # 'TYPE' column is handled separately by hardcoding 'Sale' or 'Return'
    }

    # --- Processing Function for a single file ---
    def process_file(uploaded_file, data_type):
        if uploaded_file is None:
            return None

        # Read the uploaded Excel file
        df = pd.read_excel(uploaded_file)
        
        # 1. Apply the column renaming
        df_processed = df.rename(columns=COLUMN_MAPPING)

        # 2. Select only the required columns (and ensure they exist)
        required_cols = list(COLUMN_MAPPING.values())
        
        # Check if the renamed columns exist after the mapping.
        # This handles cases where a column might be missing from the input file.
        missing_cols = [col for col in required_cols if col not in df_processed.columns]
        if missing_cols:
            st.warning(f"⚠️ Input file '{uploaded_file.name}' is missing the following columns for mapping: {missing_cols}. Please check your file headers.")
            # Drop the missing columns from the required list to proceed with available data
            required_cols = [col for col in required_cols if col not in missing_cols]

        df_final = df_processed[required_cols].copy()

        # 3. Add the 'TYPE' column (Sale/Return)
        df_final.loc[:, 'TYPE'] = data_type
        
        # 4. Handle taxable value for Returns (make it negative)
        if data_type == 'Return':
            st.info(f"Applying negative sign to 'tcs_taxable_amount' for **{data_type}** data.")
            # Ensure the column is numeric before applying the negative sign
            df_final.loc[:, 'tcs_taxable_amount'] = pd.to_numeric(
                df_final['tcs_taxable_amount'], errors='coerce'
            ).abs() * -1
        else:
            # Ensure sales values are positive
            df_final.loc[:, 'tcs_taxable_amount'] = pd.to_numeric(
                df_final['tcs_taxable_amount'], errors='coerce'
            ).abs()
        
        # Re-order the columns to match the request order
        final_order = ['order_date', 'order_num', 'hsn_code', 'gst_rate', 
                       'tcs_taxable_amount', 'end_customer_state_new', 'TYPE', 'QTY']
        # Filter the final_order to only include columns that were successfully created
        final_order_present = [col for col in final_order if col in df_final.columns]
        
        return df_final[final_order_present]


    # Process Sales Data (Tcs_sales)
    df_sales = process_file(sales_file, 'Sale')

    # Process Returns Data (Tcs_sales_return)
    df_returns = process_file(returns_file, 'Return')

    # --- Creating the Multi-Sheet Excel File ---
    if df_sales is not None or df_returns is not None:
        
        # Create the merged DataFrame (Sheet 3)
        merged_dfs = []
        if df_sales is not None:
            merged_dfs.append(df_sales)
        if df_returns is not None:
            merged_dfs.append(df_returns)
            
        if merged_dfs:
            df_merged = pd.concat(merged_dfs, ignore_index=True)
            
            # Use io.BytesIO to write the Excel file to memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                
                # Sheet 1: Sales Data
                if df_sales is not None:
                    df_sales.to_excel(writer, sheet_name='Tcs_sales', index=False)
                    st.success("✅ 'Tcs_sales' sheet generated successfully.")

                # Sheet 2: Returns Data
                if df_returns is not None:
                    df_returns.to_excel(writer, sheet_name='Tcs_sales_return', index=False)
                    st.success("✅ 'Tcs_sales_return' sheet generated successfully.")

                # Sheet 3: Merged Data
                df_merged.to_excel(writer, sheet_name='Merged_Data', index=False)
                st.success("✅ 'Merged_Data' sheet generated successfully.")
                
            # Prepare file for download
            processed_data = output.getvalue()
            return processed_data
            
    return None

# ==============================================================================
# Streamlit UI
# ==============================================================================
st.set_page_config(
    page_title="TCS Data Processor",
    layout="centered",
    initial_sidebar_state="auto"
)

st.title("📊 TCS Sales Data Consolidation Tool")
st.markdown("Upload your two Excel files to process the data, apply return logic, and create a merged, three-sheet report.")
st.markdown("---")

# 1. File Upload Widgets
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Sales Data (tcs_sales)")
    sales_file = st.file_uploader(
        "Upload the TCS Sales Excel File",
        type=['xlsx', 'xls'],
        key='sales'
    )

with col2:
    st.subheader("2. Sales Return Data (tcs_sales_return)")
    returns_file = st.file_uploader(
        "Upload the TCS Sales Return Excel File",
        type=['xlsx', 'xls'],
        key='returns'
    )

st.markdown("---")

# 2. Processing and Download
if sales_file and returns_file:
    st.subheader("3. Process and Download")
    
    if st.button("🚀 Generate Final Report"):
        with st.spinner('Processing data and creating Excel file...'):
            processed_excel_data = process_and_combine_data(sales_file, returns_file)

        if processed_excel_data:
            st.download_button(
                label="⬇️ Download Consolidated TCS Report.xlsx",
                data=processed_excel_data,
                file_name="Consolidated_TCS_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.balloons()
        else:
            st.error("❌ Failed to process data. Please ensure both files are uploaded and contain valid data.")

# 3. Next Steps/Guidance
st.sidebar.markdown("## 📚 Guidance")
st.sidebar.markdown("""
This tool maps your input data to the required output schema:
* `tcs_taxable_amount` is made **negative** for **Returns**.
* The `TYPE` column is set to **Sale** or **Return**.
* The output file contains three sheets: **Tcs_sales**, **Tcs_sales_return**, and **Merged_Data**.
""")
st.sidebar.markdown("---")
st.sidebar.warning("**To Run:** Save the code as `data_processor.py` and run `streamlit run data_processor.py` in your terminal.")
