import streamlit as st
import pandas as pd
import io
import zipfile
import requests
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ====================================================================
#  GITHUB TEMPLATE URL
# ====================================================================
GITHUB_TEMPLATE_URL = "https://raw.githubusercontent.com/Biswa-hack/Messo_GST/main/MESSO%20GST%20Template.xlsx"

# ====================================================================
#  GLOBAL COLUMN MAPPING (source_col -> target_col)
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

# exact order written into Excel columns B:I (8 columns)
WRITE_COL_ORDER = [
    'order_date',             # B
    'order_num',              # C
    'hsn_code',               # D
    'gst_rate',               # E
    'tcs_taxable_amount',     # F
    'end_customer_state_new', # G
    'QTY',                    # H
    'TYPE'                    # I
]

# ====================================================================
#  STATE → GST CODE MAPPING
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
    "Meghalaya": "17-Meghalaya",
    "Megalaya": "17-Meghalaya",
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
    "Pondicherry": "34-Puducherry",
    "Andaman and Nico.In.": "35-Andaman & Nicobar Islands",
    "ANDAMAN AND NICOBAR ISLANDS": "35-Andaman & Nicobar Islands",
    "Andaman & Nicobar Islands": "35-Andaman & Nicobar Islands",
    "Telangana": "36-Telangana",
    "Andhra Pradesh": "37-Andhra Pradesh",
    "Ladakh": "38-Ladakh",
    "Other Territory": "97-Other Territory"
}

def _normalized_mapping(mapping):
    norm = {}
    for k, v in mapping.items():
        key = str(k).strip().lower().replace("\u00A0", " ")
        key = " ".join(key.split())
        norm[key] = v
    return norm

NORM_STATE_MAPPING = _normalized_mapping(STATE_MAPPING)

# ====================================================================
#  LOAD TEMPLATE FROM GITHUB
# ====================================================================
def load_template_from_github():
    r = requests.get(GITHUB_TEMPLATE_URL)
    if r.status_code != 200:
        st.error(f"❌ Could not download template from GitHub. HTTP {r.status_code}")
        return None
    return io.BytesIO(r.content)

# ====================================================================
#  PROCESS SINGLE EXCEL FILE (Sales or Returns)
# ====================================================================
def process_file(file_data, data_type):
    df = pd.read_excel(file_data)
    df_processed = df.rename(columns=COLUMN_MAPPING)

    # Ensure all target cols exist
    for tgt in COLUMN_MAPPING.values():
        if tgt not in df_processed.columns:
            df_processed[tgt] = pd.NA

    # Build ordered dataframe (without TYPE yet)
    pre_type_cols = WRITE_COL_ORDER[:-1]  # first 7 columns (TYPE excluded)
    df_final = df_processed[pre_type_cols].copy()

    # safe numeric conversion
    df_final['tcs_taxable_amount'] = pd.to_numeric(df_final.get('tcs_taxable_amount', 0), errors='coerce').fillna(0)
    df_final['QTY'] = pd.to_numeric(df_final.get('QTY', 0), errors='coerce').fillna(0)

    # add TYPE as explicit final column (ensures position)
    df_final['TYPE'] = data_type

    # returns => negative numeric values
    if str(data_type).lower().startswith("return"):
        df_final['tcs_taxable_amount'] = df_final['tcs_taxable_amount'].abs() * -1
        df_final['QTY'] = df_final['QTY'].abs() * -1
    else:
        df_final['tcs_taxable_amount'] = df_final['tcs_taxable_amount'].abs()
        df_final['QTY'] = df_final['QTY'].abs()

    return df_final

# ====================================================================
#  PROCESS ZIP + MERGE + WRITE TO TEMPLATE (with formulas K->O)
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

    if sales_data is None or return_data is None:
        st.error("❌ Sales and Returns files not found in ZIP. Ensure both files are present and named clearly (contain 'return' for returns).")
        return None

    df_sales = process_file(sales_data, "Sale")
    df_returns = process_file(return_data, "Return")
    df_merged = pd.concat([df_sales, df_returns], ignore_index=True)

    # normalize and map states
    def _normalize_state(s):
        s = "" if pd.isna(s) else str(s)
        s = s.strip().lower().replace("\u00A0", " ")
        return " ".join(s.split())

    df_merged['end_customer_state_norm'] = df_merged['end_customer_state_new'].apply(_normalize_state)
    df_merged['J_mapped'] = df_merged['end_customer_state_norm'].map(NORM_STATE_MAPPING).fillna("")

    unmapped = (
        df_merged.loc[
            (df_merged['end_customer_state_norm'] != "") & (df_merged['J_mapped'] == ""),
            'end_customer_state_new'
        ].astype(str).unique().tolist()
    )

    # Load workbook and sheet
    template_stream = load_template_from_github()
    if template_stream is None:
        return None

    wb = load_workbook(template_stream)
    if 'raw' not in wb.sheetnames:
        st.error("❌ Template workbook does not contain a sheet named 'raw'.")
        return None
    ws = wb['raw']

    # Clear previous B:I (rows from row 3)
    for row in range(3, ws.max_row + 1):
        for col in range(2, 10):  # B (2) .. I (9)
            ws.cell(row=row, column=col).value = None

    start_row = 3

    # Ensure write columns exist in df_merged
    for col in WRITE_COL_ORDER:
        if col not in df_merged.columns:
            df_merged[col] = "" if col not in ('QTY', 'tcs_taxable_amount') else 0

    # Build the write DataFrame in exact order (guarantees TYPE & QTY positions)
    write_df = df_merged[WRITE_COL_ORDER]

    # Write columns B:I
    for r_idx, row in enumerate(dataframe_to_rows(write_df, index=False, header=False)):
        for c_idx, value in enumerate(row):
            ws.cell(start_row + r_idx, 2 + c_idx).value = value

    # Column A = "Messo"
    for r in range(len(write_df)):
        ws.cell(start_row + r, 1).value = "Messo"

    # Column J (10) = mapped state code/name
    for r in range(len(write_df)):
        ws.cell(start_row + r, 10).value = df_merged.loc[r, "J_mapped"]

    # Insert formulas into K (11) -> O (15)
    # K, L, M, N, O formulas respectively:
    FORMULAS_K_TO_O = [
        '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)',  # K
        '=IF(J{0}=$X$22,F{0}*E{0}/100/2,0)',  # L
        '=IF(J{0}=$X$22,0,F{0}*E{0}/100)',     # M
        '=K{0}+L{0}+M{0}+F{0}',                # N
        '=(M{0}+L{0}+K{0})/F{0}'               # O
    ]

    for r in range(len(write_df)):
        excel_row = start_row + r
        for offset, formula in enumerate(FORMULAS_K_TO_O):
            ws.cell(row=excel_row, column=11 + offset).value = formula.format(excel_row)

    # Save workbook to bytes
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue(), unmapped

# ====================================================================
#  STREAMLIT UI
# ====================================================================
st.set_page_config(page_title="TCS Processor", layout="wide")
st.title("📊 TCS Data Integration & Template Filler")
st.markdown("---")

zipped_files = st.file_uploader("Upload ZIP containing Sales & Returns Excel files", type=["zip"])

if zipped_files:
    if st.button("🚀 Generate Report"):
        with st.spinner("Processing..."):
            result = process_zip_and_combine_data(zipped_files)

        if result:
            file_bytes, unmapped_states = result
            st.download_button(
                "⬇ Download Modified_Combo_Report.xlsx",
                file_bytes,
                "Modified_Combo_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.success("Done!")

            if unmapped_states:
                st.warning("Some state names were not mapped. Add variants to STATE_MAPPING if needed.")
                st.write("Unmapped (unique):")
                st.write(unmapped_states)
            else:
                st.info("All state values mapped successfully.")
