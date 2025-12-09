import streamlit as st
import pandas as pd
import io
import zipfile
import requests
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# ====================================================================
#  GITHUB TEMPLATE URL (UPDATED)
# ====================================================================
GITHUB_TEMPLATE_URL = (
    "https://raw.githubusercontent.com/Biswa-hack/Messo_GST/main/MESSO%20GST%20Template.xlsx"
)


# ====================================================================
#  GLOBAL COLUMN MAPPING
# ====================================================================
COLUMN_MAPPING = {
    'order_date': 'order_date',
    'sub_order_num': 'order_num',
    'hsn_code': 'hsn_code',
    'gst_rate': 'gst_rate',
    'total_taxable_sale_value': 'tcs_taxable_amount',
    'end_customer_state_new': 'end_customer_state_new',
    'quantity': 'QTY',
}

# ====================================================================
#  STATE → GST CODE MAPPING (Column J)
# ====================================================================
STATE_MAPPING = {
    "Jammu and Kashmir": "01-Jammu & Kashmir",
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
    "MEGHALAYA": "17-Meghalaya",
    "Assam": "18-Assam",
    "West Bengal": "19-West Bengal",
    "Jharkhand": "20-Jharkhand",
    "Odisha": "21-Odisha",
    "Chhattisgarh": "22-Chhattisgarh",
    "Madhya Pradesh": "23-Madhya Pradesh",
    "Gujarat": "24-Gujarat",
    "Daman and Diu": "25-Daman & Diu",
    "Daman & Diu": "25-Daman & Diu",
    "THE DADRA AND NAGAR HAVELI AND DAMAN AND DIU": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Dadra & Nagar Haveli & Daman & Diu": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Dadra and Nagar Haveli": "26-Dadra & Nagar Haveli & Daman & Diu",
    "Maharashtra": "27-Maharashtra",
    "Karnataka": "29-Karnataka",
    "Goa": "30-Goa",
    "Lakshadweep": "31-Lakshdweep",
    "Kerala": "32-Kerala",
    "Tamil Nadu": "33-Tamil Nadu",
    "PONDICHERRY": "34-Puducherry",
    "Puducherry": "34-Puducherry",
    "Andaman and Nico.In.": "35-Andaman & Nicobar Islands",
    "ANDAMAN AND NICOBAR ISLANDS": "35-Andaman & Nicobar Islands",
    "Andaman & Nicobar Islands": "35-Andaman & Nicobar Islands",
    "Telangana": "36-Telangana",
    "Andhra Pradesh": "37-Andhra Pradesh",
    "Ladakh": "38-Ladakh",
    "Other Territory": "97-Other Territory"
}


# ====================================================================
#  LOAD TEMPLATE FROM GITHUB URL
# ====================================================================
def load_template_from_github():
    r = requests.get(GITHUB_TEMPLATE_URL)
    if r.status_code != 200:
        st.error("❌ Could not download template from GitHub.")
        return None
    return io.BytesIO(r.content)


# ====================================================================
#  PROCESS SINGLE EXCEL FILE (Sales or Returns)
# ====================================================================
def process_file(file_data, data_type):
    df = pd.read_excel(file_data)
    df_processed = df.rename(columns=COLUMN_MAPPING)

    result_cols = list(COLUMN_MAPPING.values())
    df_final = df_processed[result_cols].copy()
    df_final["TYPE"] = data_type

    # Negative values for returns
    if data_type == "Return":
        df_final["tcs_taxable_amount"] = pd.to_numeric(
            df_final["tcs_taxable_amount"], errors="coerce"
        ).abs() * -1

        df_final["QTY"] = pd.to_numeric(
            df_final["QTY"], errors="coerce"
        ).abs() * -1

    # Positive for sales
    else:
        df_final["tcs_taxable_amount"] = pd.to_numeric(
            df_final["tcs_taxable_amount"], errors="coerce"
        ).abs()

        df_final["QTY"] = pd.to_numeric(
            df_final["QTY"], errors="coerce"
        ).abs()

    return df_final


# ====================================================================
#  PROCESS ZIP + MERGE + WRITE TO TEMPLATE
# ====================================================================
def process_zip_and_combine_data(zip_file):

    sales_data = None
    return_data = None

    # unzip
    with zipfile.ZipFile(io.BytesIO(zip_file.read())) as z:
        for name in z.namelist():
            if name.endswith((".xlsx", ".xls")):
                if "return" in name.lower() or "rtn" in name.lower():
                    return_data = z.open(name)
                else:
                    sales_data = z.open(name)

    if not sales_data or not return_data:
        st.error("❌ Sales and Returns files not found in ZIP.")
        return None

    df_sales = process_file(sales_data, "Sale")
    df_returns = process_file(return_data, "Return")

    df_merged = pd.concat([df_sales, df_returns], ignore_index=True)

    # Create J column using Python mapping
    df_merged["J_mapped"] = df_merged["end_customer_state_new"].map(STATE_MAPPING).fillna("")

    # Load template from GitHub
    template_stream = load_template_from_github()
    if template_stream is None:
        return None

    wb = load_workbook(template_stream)
    ws = wb["raw"]

    # Clear B:I
    for row in range(3, ws.max_row + 1):
        for col in range(2, 10):
            ws.cell(row=row, column=col).value = None

    start_row = 3

    # Insert B:I
    for r_idx, row in enumerate(
        dataframe_to_rows(df_merged.iloc[:, :8], index=False, header=False)
    ):
        for c_idx, value in enumerate(row):
            ws.cell(start_row + r_idx, 2 + c_idx).value = value

    # Insert Column A = "Messo"
    for r in range(len(df_merged)):
        ws.cell(start_row + r, 1).value = "Messo"

    # Insert Column J (mapped)
    for r in range(len(df_merged)):
        ws.cell(start_row + r, 10).value = df_merged.loc[r, "J_mapped"]

    # Insert formulas for K–O
    RAW_FORMULAS = [
        None,
        '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)',
        '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)',
        '=IF(J{0}=$X$22,0,F{0}*E{0}/100)',
        '=K{0}+L{0}+M{0}+F{0}',
        '=(M{0}+L{0}+K{0})/F{0}'
    ]

    for r in range(len(df_merged)):
        excel_row = start_row + r
        for col_offset, formula in enumerate(RAW_FORMULAS):
            if formula is not None:
                ws.cell(
                    excel_row,
                    10 + col_offset,
                    formula.format(excel_row)
                )

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ====================================================================
#  STREAMLIT UI
# ====================================================================
st.set_page_config(page_title="TCS Processor", layout="wide")

st.title("📊 TCS Data Integration & Template Filler (GitHub Template Auto-loaded)")
st.markdown("---")

zipped_files = st.file_uploader("Upload ZIP containing Sales & Returns Excel files", type=["zip"])

if zipped_files:
    if st.button("🚀 Generate Report"):
        with st.spinner("Processing..."):
            result = process_zip_and_combine_data(zipped_files)

        if result:
            st.download_button(
                "⬇ Download Modified_Combo_Report.xlsx",
                result,
                "Modified_Combo_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Done!")
