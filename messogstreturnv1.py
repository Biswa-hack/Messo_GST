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
# IMPORTANT: Update this to your supplier's GST state code (e.g., Maharashtra)
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
    "Jammu And Kashmir": "01-Jammu & Kashmir", "Jammu & Kashmir": "01-Jammu & Kashmir",
    "Himachal Pradesh": "02-Himachal Pradesh", "Punjab": "03-Punjab",
    "Chandigarh": "04-Chandigarh", "Uttarakhand": "05-Uttarakhand",
    "Haryana": "06-Haryana", "Delhi": "07-Delhi",
    "Rajasthan": "08-Rajasthan", "Uttar Pradesh": "09-Uttar Pradesh",
    "Bihar": "10-Bihar", "Sikkim": "11-Sikkim",
    "Arunachal Pradesh": "12-Arunachal Pradesh", "Nagaland": "13-Nagaland",
    "Manipur": "14-Manipur", "Mizoram": "15-Mizoram",
    "Tripura": "16-Tripura", "Megalaya": "17-Meghalaya",
    "Meghalaya": "17-Meghalaya", "Assam": "18-Assam",
    "West Bengal": "19-West Bengal", "Jharkhand": "20-Jharkhand",
    "Odisha": "21-Odisha", "Chhattisgarh": "22-Chhattisgarh",
    "Madhya Pradesh": "23-Madhya Pradesh", "Gujarat": "24-Gujarat",
    "Daman And Diu": "25-Daman & Diu", "Daman & Diu": "25-Daman & Diu",
    "The Dadra And Nagar Haveli And Daman And Diu": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Dadra & Nagar Haveli & Daman & Diu": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Dadra And Nagar Haveli": "26-Dadra & Nagar Haveli & Daman & Diu", 
    "Maharashtra": "27-Maharashtra", "Karnataka": "29-Karnataka",
    "Goa": "30-Goa", "Lakshadweep": "31-Lakshdweep",
    "Kerala": "32-Kerala", "Tamil Nadu": "33-Tamil Nadu",
    "Pondicherry": "34-Puducherry", "Puducherry": "34-Puducherry",
    "Andaman And Nico.In.": "35-Andaman & Nicobar Islands", 
    "Andaman And Nicobar Islands": "35-Andaman & Nicobar Islands", 
    "Andaman & Nicobar Islands": "35-Andaman & Nicobar Islands",
    "Telangana": "36-Telangana", "Andhra Pradesh": "37-Andhra Pradesh",
    "Ladakh": "38-Ladakh", "Other Territory": "97-Other Territory"
}

# ============================================================
#  HELPER FUNCTIONS
# ============================================================
def load_template_from_github():
    """Downloads the Excel template from the specified GitHub URL."""
    r = requests.get(GITHUB_TEMPLATE_URL)
    if r.status_code != 200:
        st.error("❌ Could not download template from GitHub. Status Code: " + str(r.status_code))
        return None
    return io.BytesIO(r.content)

def calculate_tax_components(df):
    """
    Calculates CGST, SGST, IGST, and Total Tax for the DataFrame based on 
    the Place of Supply (J_mapped) vs. the DEFAULT_SUPPLIER_STATE_CODE.
    """
    # Create a copy to perform calculations without warning
    df_taxed = df.copy() 
    
    is_intra_state = df_taxed["J_mapped"] == DEFAULT_SUPPLIER_STATE_CODE
    
    # Calculate Total Tax Amount (IGST Rate = GST Rate)
    # Handle division by zero/invalid rate by coercing to numeric and filling NaN
    df_taxed["gst_rate"] = pd.to_numeric(df_taxed["gst_rate"], errors='coerce').fillna(0)
    
    total_tax_rate = df_taxed["gst_rate"] / 100
    total_tax = df_taxed["tcs_taxable_amount"] * total_tax_rate
    
    # Calculate CGST/SGST (for Intra-State) and IGST (for Inter-State)
    df_taxed["CGST"] = total_tax.where(is_intra_state, 0) / 2
    df_taxed["SGST"] = total_tax.where(is_intra_state, 0) / 2
    df_taxed["IGST"] = total_tax.where(~is_intra_state, 0)
    df_taxed["Total Tax"] = df_taxed["CGST"] + df_taxed["SGST"] + df_taxed["IGST"]
    df_taxed["Total Value"] = df_taxed["tcs_taxable_amount"] + df_taxed["Total Tax"]
    
    return df_taxed

# ============================================================
#  SUMMARY GENERATION FUNCTIONS
# ============================================================

def generate_b2cs_csv(df_merged_taxed):
    """Generates the GSTR-1 B2CS (Table 7) summary in CSV format."""
    
    # Group by Place of Supply and Rate, summing Taxable Value
    summary_df = df_merged_taxed.groupby(["J_mapped", "gst_rate"]).agg(
        Taxable_Value=('tcs_taxable_amount', 'sum')
    ).reset_index()

    # Apply the required format and fixed values
    summary_df['Type'] = 'OE' # Other than E-Commerce
    summary_df['Place Of Supply'] = summary_df['J_mapped']
    summary_df['Rate'] = summary_df['gst_rate']
    summary_df['Applicable % of Tax Rate'] = '' 
    summary_df['Cess Amount'] = ''
    summary_df['E-Commerce GSTIN'] = ''
    
    # Select and reorder columns
    final_b2cs_df = summary_df[[
        'Type', 
        'Place Of Supply', 
        'Rate', 
        'Applicable % of Tax Rate', 
        'Taxable_Value', # This will be renamed later
        'Cess Amount', 
        'E-Commerce GSTIN'
    ]].rename(columns={'Taxable_Value': 'Taxable Value'})
    
    # Convert to CSV
    csv_output = final_b2cs_df.to_csv(index=False).encode('utf-8')
    return csv_output

def generate_hsn_summary(df_merged_taxed):
    """Generates the GSTR-1 HSN Summary (Table 12) in Excel format."""

    # Group by HSN and Rate, summing all relevant values
    summary_df = df_merged_taxed.groupby(["hsn_code", "gst_rate"]).agg(
        Total_Quantity=('QTY', 'sum'),
        Total_Taxable_Value=('tcs_taxable_amount', 'sum'),
        Total_Value=('Total Value', 'sum'), # Gross Value including tax
        Integrated_Tax_Amount=('IGST', 'sum'),
        Central_Tax_Amount=('CGST', 'sum'),
        State_UT_Tax_Amount=('SGST', 'sum')
    ).reset_index()

    # Apply the required format and fixed values
    summary_df['Description'] = '' # User requested constant/blank
    summary_df['UQC'] = 'NOS' # NOS-NUMBERS is common, simplifying to NOS
    summary_df['Cess Amount'] = 0.0 # Standard zero for numeric column

    # Select and reorder columns
    final_hsn_df = summary_df[[
        'hsn_code',
        'Description',
        'UQC',
        'Total_Quantity',
        'Total_Value',
        'Total_Taxable_Value',
        'Integrated_Tax_Amount',
        'Central_Tax_Amount',
        'State_UT_Tax_Amount',
        'Cess Amount',
        'gst_rate'
    ]].rename(columns={
        'hsn_code': 'HSN',
        'Total_Quantity': 'Total Quantity',
        'Total_Value': 'Total Value',
        'Total_Taxable_Value': 'Taxable Value',
        'Integrated_Tax_Amount': 'Integrated Tax Amount',
        'Central_Tax_Amount': 'Central Tax Amount',
        'State_UT_Tax_Amount': 'State/UT Tax Amount',
        'gst_rate': 'Rate'
    })
    
    # Convert to Excel
    excel_output = io.BytesIO()
    final_hsn_df.to_excel(excel_output, index=False, sheet_name='HSN_Summary')
    return excel_output.getvalue()

# ============================================================
#  MAIN ZIP PROCESSOR
# ============================================================
def process_zip_and_combine_data(zip_file):
    """Extracts, processes, merges data, fills the Excel template, and generates summaries."""
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
        return None, None, None
        
    if not sales_data or not return_data:
        st.error("❌ ZIP must contain both a **Sales** and a **Return** file.")
        return None, None, None

    # 2. Process and Merge DataFrames
    try:
        df_sales = process_file(sales_data, "Sale")
        df_returns = process_file(return_data, "Return")
    except Exception as e:
        st.error(f"❌ Error processing input files: {e}")
        return None, None, None
        
    df_merged = pd.concat([df_sales, df_returns], ignore_index=True)

    # Standardize State Names
    df_merged["end_customer_state_new"] = df_merged["end_customer_state_new"].str.title()
    
    # Map State Code (Column J)
    df_merged["J_mapped"] = df_merged["end_customer_state_new"].map(STATE_MAPPING).fillna("")

    # 3. Calculate Tax Components for Summaries and Raw Data
    df_merged_taxed = calculate_tax_components(df_merged.copy()) 

    # 4. Generate Summary Reports
    b2cs_csv_output = generate_b2cs_csv(df_merged_taxed)
    hsn_excel_output = generate_hsn_summary(df_merged_taxed)

    # 5. Load Template and Insert Raw Data
    template_stream = load_template_from_github()
    if template_stream is None:
        return None, None, None

    wb = load_workbook(template_stream)
    ws = wb["raw"]

    # Clear old data
    for row in range(3, ws.max_row + 1):
        for col in range(1, 16):
            ws.cell(row=row, column=col).value = None

    start_row = 3
    num_rows = len(df_merged)

    # Insert B → I
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
        # NOTE: $X$22 is assumed to contain the default supplier state code in the template
        # K: CGST 
        ws.cell(excel_row, 11).value = f"=IF(J{excel_row}=$X$22,F{excel_row}*E{excel_row}/100/2,0)"
        # L: SGST 
        ws.cell(excel_row, 12).value = f"=IF(J{excel_row}=$X$22,F{excel_row}*E{excel_row}/100/2,0)"
        # M: IGST 
        ws.cell(excel_row, 13).value = f"=IF(J{excel_row}<>$X$22,F{excel_row}*E{excel_row}/100,0)"
        # N: Total Value
        ws.cell(excel_row, 14).value = f"=K{excel_row}+L{excel_row}+M{excel_row}+F{excel_row}"
        # O: % GST
        ws.cell(excel_row, 15).value = f"=(K{excel_row}+L{excel_row}+M{excel_excel_row})/F{excel_row}" # Note: corrected typo on L1

    template_output = io.BytesIO()
    wb.save(template_output)
    
    # Return all three outputs
    return template_output.getvalue(), b2cs_csv_output, hsn_excel_output

# ============================================================
#  STREAMLIT UI
# ============================================================
st.set_page_config(page_title="TCS Processor V2.1", layout="wide")
st.title("📊 TCS Data Integration & GSTR-1 Summary Generator")
st.markdown("---")
st.info(f"**Default Supplier State Code (for CGST/SGST determination):** `{DEFAULT_SUPPLIER_STATE_CODE}`. Update the constant in the code if needed.")
st.markdown("---")

zipped_files = st.file_uploader("Upload ZIP containing Sales + Return files", type=["zip"])

if zipped_files:
    if st.button("🚀 Generate All 3 Reports"):
        with st.spinner("Processing... Generating Combo, B2CS Summary (CSV), and HSN Summary (Excel)."):
            combo_result, b2cs_result, hsn_result = process_zip_and_combine_data(zipped_files)

        if combo_result and b2cs_result and hsn_result:
            st.success("✔️ Done! All three reports are ready for download.")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Download 1: The original combo file
                st.markdown("#### 1. Raw Combo Data")
                st.download_button(
                    "⬇ Combo Report (.xlsx)",
                    combo_result,
                    "Modified_Combo_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                # Download 2: The new B2CS summary file (CSV)
                st.markdown("#### 2. B2CS Summary (GSTR-1 Table 7)")
                st.download_button(
                    "⬇ B2CS Summary (.csv)",
                    b2cs_result,
                    "B2CS_Summary_Report.csv",
                    mime="text/csv"
                )

            with col3:
                # Download 3: The new HSN summary file (Excel)
                st.markdown("#### 3. HSN Summary (GSTR-1 Table 12)")
                st.download_button(
                    "⬇ HSN Summary (.xlsx)",
                    hsn_result,
                    "HSN_Summary_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
