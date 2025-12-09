import streamlit as st
import pandas as pd
import io
import zipfile
import requests
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================================================
#  GITHUB RAW TEMPLATE URL
# ============================================================
GITHUB_TEMPLATE_URL = "https://raw.githubusercontent.com/Biswa-hack/Messo_GST/main/MESSO%20GST%20Template.xlsx"

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

# ============================================================
#  LOAD TEMPLATE
# ============================================================
def load_template_from_github():
    r = requests.get(GITHUB_TEMPLATE_URL)
    if r.status_code != 200:
        st.error("❌ Could not download template from GitHub.")
        return None
    return io.BytesIO(r.content)

# ============================================================
#  PROCESS INDIVIDUAL FILE
# ============================================================
def process_file(file_data, data_type):
    df = pd.read_excel(file_data)
    df_processed = df.rename(columns=COLUMN_MAPPING)

    df_final = df_processed[list(COLUMN_MAPPING.values())].copy()
    df_final["TYPE"] = data_type

    df_final["tcs_taxable_amount"] = pd.to_numeric(df_final["tcs_taxable_amount"], errors="coerce")
    df_final["QTY"] = pd.to_numeric(df_final["QTY"], errors="coerce")

    if data_type == "Return":
        df_final["tcs_taxable_amount"] = df_final["tcs_taxable_amount"].abs() * -1
        df_final["QTY"] = df_final["QTY"].abs() * -1
    else:
        df_final["tcs_taxable_amount"] = df_final["tcs_taxable_amount"].abs()
        df_final["QTY"] = df_final["QTY"].abs()

    return df_final

# ============================================================
#  MAIN ZIP PROCESSOR
# ============================================================
def process_zip_and_combine_data(zip_file):

    sales_data = None
    return_data = None

    with zipfile.ZipFile(io.BytesIO(zip_file.read())) as z:
        for name in z.namelist():
            if name.endswith((".xlsx", ".xls")):
                if "return" in name.lower() or "rtn" in name.lower():
                    return_data = z.open(name)
                else:
                    sales_data = z.open(name)

    if not sales_data or not return_data:
        st.error("❌ ZIP must contain both Sales & Return files.")
        return None

    df_sales = process_file(sales_data, "Sale")
    df_returns = process_file(return_data, "Return")

    df_merged = pd.concat([df_sales, df_returns], ignore_index=True)

    # Add State Code column J
    df_merged["J_mapped"] = df_merged["end_customer_state_new"].map(STATE_MAPPING).fillna("")

    # Load Template
    template_stream = load_template_from_github()
    if template_stream is None:
        return None

    wb = load_workbook(template_stream)
    ws = wb["raw"]

    # Clear old data
    for row in range(3, ws.max_row + 1):
        for col in range(1, 16):
            ws.cell(row=row, column=col).value = None

    start_row = 3

    # Insert B → I
    write_df = df_merged[WRITE_COL_ORDER]

    for r_idx, row in enumerate(dataframe_to_rows(write_df, index=False, header=False)):
        for c_idx, value in enumerate(row):
            ws.cell(start_row + r_idx, 2 + c_idx).value = value

    # Insert Column A = Messo
    for r in range(len(df_merged)):
        ws.cell(start_row + r, 1).value = "Messo"

    # Insert Column J
    for r in range(len(df_merged)):
        ws.cell(start_row + r, 10).value = df_merged.loc[r, "J_mapped"]

    # Insert formulas K–O
    for r in range(len(df_merged)):
        excel_row = start_row + r
        ws.cell(excel_row, 11).value = f"=IF(J{excel_row}=$X$22,F{excel_row}*E{excel_row}/100/2,0)"
        ws.cell(excel_row, 12).value = f"=IF(J{excel_row}=$X$22,F{excel_row}*E{excel_row}/100/2,0)"
        ws.cell(excel_row, 13).value = f"=IF(J{excel_row}<>$X$22,F{excel_row}*E{excel_row}/100,0)"
        ws.cell(excel_row, 14).value = f"=K{excel_row}+L{excel_row}+M{excel_row}+F{excel_row}"
        ws.cell(excel_row, 15).value = f"=(K{excel_row}+L{excel_row}+M{excel_row})/F{excel_row}"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================================================
#  STREAMLIT UI
# ============================================================
st.set_page_config(page_title="TCS Processor", layout="wide")
st.title("📊 TCS Data Integration & Template Filler")
st.markdown("---")

zipped_files = st.file_uploader("Upload ZIP containing Sales + Return files", type=["zip"])

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
            st.success("✔️ Done!")
