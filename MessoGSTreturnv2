import streamlit as st
import pandas as pd
import io
import zipfile
import requests
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================================================
#  GLOBAL CONSTANTS
# ============================================================
GITHUB_TEMPLATE_URL = "https://raw.githubusercontent.com/Biswa-hack/Messo_GST/main/MESSO%20GST%20Template.xlsx"
# Assuming the default supplier state for Intra-state calculation is Maharashtra (Code 27)
DEFAULT_SUPPLIER_STATE_CODE = "27-Maharashtra" 

# ============================================================
#  COLUMN MAPPING
# ============================================================
COLUMN_MAPPING = {
    'order_date': 'order_date',
    'sub_order_num': 'order_num',
    'hsn_code': 'hsn_code',
    'gst_rate': 'gst_rate',
    'total_taxable_sale_value': 'tcs_taxable_amount',
    'end_customer_state_new': 'end_customer_state_new',
    'quantity': 'QTY'
}

# ============================================================
#  OUTPUT COLUMN ORDER (B → I)
# ============================================================
WRITE_COL_ORDER = [
    'order_date',             # B
    'order_num',              # C
    'hsn_code',               # D
    'gst_rate',               # E
    'tcs_taxable_amount',     # F
    'end_customer_state_new', # G
    'TYPE',                   # H
    'QTY'                     # I
]

# ============================================================
#  STATE → GST CODE MAPPING (Column J)
# ============================================================
STATE_MAPPING = {
    "Jammu And Kashmir": "01-Jammu & Kashmir", 
    "Jammu & Kashmir": "01-Jammu & Kashmir",
    "Himachal Pradesh": "02-Himachal Pradesh",
    "Punjab": "03-Punjab",
    "Chandigarh": "04-Chandigarh",
    "Uttarakhand": "05-Uttarakhand",
    "Haryana": "06-Haryana",
    "Delhi": "07-Delhi",
    "Rajasthan": "08-Rajasthan",
    "Uttar Pradesh": "09-Uttar Pradesh",
    "Bihar": "10-Bihar",
    "Sikkim": "11-Sikkim",
    "Arunachal Pradesh": "12-Arunachal Pradesh",
    "Nagaland": "13-Nagaland",
    "Manipur": "14-Manipur",
    "Mizoram": "15-Mizoram",
    "Tripura": "16-Tripura",
    "Megalaya": "17-Meghalaya",
    "Meghalaya": "17-Meghalaya", 
    "Assam": "18-Assam",
    "West Bengal": "19-West Bengal",
    "Jharkhand": "20-Jharkhand",
    "Odisha": "21-Odisha",
    "Chhattisgarh": "22-Chhattisgarh",
    "Madhya Pradesh": "23-Madhya Pradesh",
    "Gujarat": "24-Gujarat",
    "Daman And Diu": "25-Daman & Diu", 
    "Daman & Diu": "25-Daman & Diu",
    "The Dadra And Nagar Haveli And Daman And Diu": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Dadra & Nagar Haveli & Daman & Diu": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Dadra And Nagar Haveli": "26-Dadra & Nagar Haveli & Daman & Diu", 
    "Maharashtra": "27-Maharashtra",
    "Karnataka": "29-Karnataka",
    "Goa": "30-Goa",
    "Lakshadweep": "31-Lakshdweep",
    "Kerala": "32-Kerala",
    "Tamil Nadu": "33-Tamil Nadu",
    "Pondicherry": "34-Puducherry", 
    "Puducherry": "34-Puducherry",
    "Andaman And Nico.In.": "35-Andaman & Nicobar Islands", 
    "Andaman And Nicobar Islands": "35-Andaman & Nicobar Islands", 
    "Andaman & Nicobar Islands": "35-Andaman & Nicobar Islands",
    "Telangana": "36-Telangana",
    "Andhra Pradesh": "37-Andhra Pradesh",
    "Ladakh": "38-Ladakh",
    "Other Territory": "97-Other Territory"
}

# ============================================================
#  LOAD TEMPLATE
# ============================================================
def load_template_from_github():
    """Downloads the Excel template from the specified GitHub URL."""
    r = requests.get(GITHUB_TEMPLATE_URL)
    if r.status_code != 200:
        st.error("❌ Could not download template from GitHub. Status Code: " + str(r.status_code))
        return None
    return io.BytesIO(r.content)

# ============================================================
#  PROCESS INDIVIDUAL FILE
# ============================================================
def process_file(file_data, data_type):
    """Reads Excel, renames columns, and adjusts values for Sales/Return."""
    df = pd.read_excel(file_data)
    df_processed = df.rename(columns=COLUMN_MAPPING)

    df_final = df_processed[list(COLUMN_MAPPING.values())].copy()
    df_final["TYPE"] = data_type

    df_final["tcs_taxable_amount"] = pd.to_numeric(df_final["tcs_taxable_amount"], errors="coerce")
    df_final["QTY"] = pd.to_numeric(df_final["QTY"], errors="coerce")

    if data_type == "Return":
        # Returns should be negative (Taxable Value for returns is negative)
        df_final["tcs_taxable_amount"] = df_final["tcs_taxable_amount"].abs() * -1
        df_final["QTY"] = df_final["QTY"].abs() * -1
    else:
        # Sales should be positive (Taxable Value for sales is positive)
        df_final["tcs_taxable_amount"] = df_final["tcs_taxable_amount"].abs()
        df_final["QTY"] = df_final["QTY"].abs()

    return df_final

# ============================================================
#  SUMMARY GENERATION (NEW FUNCTION FOR VERSION 2)
# ============================================================
def generate_gstr1_summary(df_merged):
    """Generates a B2CS-style summary (Place of Supply wise) with tax calculations."""

    # 1. Calculate Tax Components for each row
    is_intra_state = df_merged["J_mapped"] == DEFAULT_SUPPLIER_STATE_CODE
    
    # Calculate Total Tax Amount
    total_tax = df_merged["tcs_taxable_amount"] * df_merged["gst_rate"] / 100
    
    # Calculate CGST/SGST (for Intra-State) and IGST (for Inter-State)
    df_merged["CGST"] = total_tax.where(is_intra_state, 0) / 2
    df_merged["SGST"] = total_tax.where(is_intra_state, 0) / 2
    df_merged["IGST"] = total_tax.where(~is_intra_state, 0)
    df_merged["Total Tax"] = df_merged["CGST"] + df_merged["SGST"] + df_merged["IGST"]
    df_merged["Total Value"] = df_merged["tcs_taxable_amount"] + df_merged["Total Tax"]

    # 2. Aggregate the data
    summary_df = df_merged.groupby(["J_mapped", "gst_rate", "TYPE"]).agg(
        Total_Taxable_Value=('tcs_taxable_amount', 'sum'),
        Total_CGST=('CGST', 'sum'),
        Total_SGST=('SGST', 'sum'),
        Total_IGST=('IGST', 'sum'),
        Total_Qty=('QTY', 'sum')
    ).reset_index()

    # 3. Rename and format columns for final summary output
    summary_df = summary_df.rename(columns={
        'J_mapped': 'Place of Supply (GST Code)',
        'gst_rate': 'GST Rate',
        'TYPE': 'Transaction Type'
    })

    # Add a Gross Value column
    summary_df['Total_Invoice_Value'] = summary_df['Total_Taxable_Value'] + summary_df['Total_CGST'] + summary_df['Total_SGST'] + summary_df['Total_IGST']

    # Reorder columns
    summary_df = summary_df[[
        'Place of Supply (GST Code)', 
        'GST Rate', 
        'Transaction Type', 
        'Total_Taxable_Value', 
        'Total_CGST', 
        'Total_SGST', 
        'Total_IGST', 
        'Total_Invoice_Value',
        'Total_Qty'
    ]]
    
    return summary_df

# ============================================================
#  MAIN ZIP PROCESSOR
# ============================================================
def process_zip_and_combine_data(zip_file):

    sales_data = None
    return_data = None

    # 1. Extract Sales and Return files from ZIP
    try:
        with zipfile.ZipFile(io.BytesIO(zip_file.read())) as z:
            for name in z.namelist():
                if name.endswith((".xlsx", ".xls")):
                    if "return" in name.lower() or "rtn" in name.lower():
                        return_data = z.open(name)
                    elif "sale" in name.lower() or "sls" in name.lower() or "invoice" in name.lower():
                        sales_data = z.open(name)
    except zipfile.BadZipFile:
        st.error("❌ Invalid or corrupted ZIP file.")
        return None
        
    if not sales_data or not return_data:
        st.error("❌ ZIP must contain both a **Sales** and a **Return** file.")
        return None

    # 2. Process and Merge DataFrames
    try:
        df_sales = process_file(sales_data, "Sale")
        df_returns = process_file(return_data, "Return")
    except Exception as e:
        st.error(f"❌ Error processing input files: {e}")
        return None
        
    df_merged = pd.concat([df_sales, df_returns], ignore_index=True)

    # Standardize State Names (Fix from previous version)
    df_merged["end_customer_state_new"] = df_merged["end_customer_state_new"].str.title()
    
    # Map State Code (Column J)
    df_merged["J_mapped"] = df_merged["end_customer_state_new"].map(STATE_MAPPING).fillna("")

    # 3. Generate Summary Report (NEW STEP)
    summary_df = generate_gstr1_summary(df_merged.copy()) # Use a copy to avoid altering the original df_merged

    # 4. Load Template
    template_stream = load_template_from_github()
    if template_stream is None:
        return None

    wb = load_workbook(template_stream)
    ws = wb["raw"]

    # 5. Clear old data from Row 3 onwards
    for row in range(3, ws.max_row + 1):
        for col in range(1, 16):
            ws.cell(row=row, column=col).value = None

    start_row = 3
    num_rows = len(df_merged)

    # 6. Insert Data into Excel Template
    write_df = df_merged[WRITE_COL_ORDER]

    for r_idx, row in enumerate(dataframe_to_rows(write_df, index=False, header=False)):
        for c_idx, value in enumerate(row):
            ws.cell(start_row + r_idx, 2 + c_idx).value = value

    # Insert Column A = Messo
    for r in range(num_rows):
        ws.cell(start_row + r, 1).value = "Messo"

    # Insert Column J (Mapped State Code)
    for r in range(num_rows):
        ws.cell(start_row + r, 10).value = df_merged.loc[r, "J_mapped"]

    # Insert formulas K–O
    for r in range(num_rows):
        excel_row = start_row + r
        # K: CGST (IF J = Default State)
        ws.cell(excel_row, 11).value = f"=IF(J{excel_row}=$X$22,F{excel_row}*E{excel_row}/100/2,0)"
        # L: SGST (IF J = Default State)
        ws.cell(excel_row, 12).value = f"=IF(J{excel_row}=$X$22,F{excel_row}*E{excel_row}/100/2,0)"
        # M: IGST (IF J != Default State)
        ws.cell(excel_row, 13).value = f"=IF(J{excel_row}<>$X$22,F{excel_row}*E{excel_row}/100,0)"
        # N: Total Value (Taxable + Total Tax)
        ws.cell(excel_row, 14).value = f"=K{excel_row}+L{excel_row}+M{excel_row}+F{excel_row}"
        # O: % GST (Rate based on Tax/Taxable)
        ws.cell(excel_row, 15).value = f"=(K{excel_row}+L{excel_row}+M{excel_row})/F{excel_row}"

    # Prepare outputs
    template_output = io.BytesIO()
    wb.save(template_output)
    
    summary_output = io.BytesIO()
    # Save the summary dataframe to an Excel file
    summary_df.to_excel(summary_output, index=False, sheet_name='B2CS_Summary')
    
    return template_output.getvalue(), summary_output.getvalue()

# ============================================================
#  STREAMLIT UI
# ============================================================
st.set_page_config(page_title="TCS Processor V2", layout="wide")
st.title("📊 TCS Data Integration & Template Filler V2 (with Summary)")
st.markdown("---")

zipped_files = st.file_uploader("Upload ZIP containing Sales + Return files", type=["zip"])

if zipped_files:
    if st.button("🚀 Generate Reports"):
        with st.spinner("Processing... Generating Combo and Summary Reports."):
            results = process_zip_and_combine_data(zipped_files)

        if results:
            template_result, summary_result = results

            # Download 1: The original combo file
            st.download_button(
                "⬇ Download 1: Modified Combo Report (Template)",
                template_result,
                "Modified_Combo_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Download 2: The new summary file
            st.download_button(
                "⬇ Download 2: B2CS Summary Report",
                summary_result,
                "Summary_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success("✔️ Done! Both the Modified Combo Report and the B2CS Summary Report are ready for download.")
