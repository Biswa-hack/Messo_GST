import streamlit as st
import pandas as pd
import io
import zipfile
import requests
import json
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="GSTR-1 Report Generator ✨",
    layout="wide",
    initial_sidebar_state="collapsed"
)

if 'combo_result' not in st.session_state:
    st.session_state.combo_result = None
    st.session_state.b2cs_result = None
    st.session_state.hsn_result = None
    st.session_state.json_result = None
    st.session_state.file_name = None
    st.session_state.dynamic_gstin = "N/A"
    st.session_state.dynamic_fp = "N/A"
    st.session_state.default_state_code_numeric = "N/A"

# ============================================================
# CONSTANTS
# ============================================================

GITHUB_TEMPLATE_URL = "https://raw.githubusercontent.com/Biswa-hack/Messo_GST/main/MESSO%20GST%20Template.xlsx"

COLUMN_MAPPING = {
    'order_date': 'order_date',
    'sub_order_num': 'order_num',
    'hsn_code': 'hsn_code',
    'gst_rate': 'gst_rate',
    'total_taxable_sale_value': 'tcs_taxable_amount',
    'end_customer_state_new': 'end_customer_state_new',
    'quantity': 'QTY'
}

WRITE_COL_ORDER = [
    'order_date', 'order_num', 'hsn_code',
    'gst_rate', 'tcs_taxable_amount',
    'end_customer_state_new', 'TYPE', 'QTY'
]

STATE_MAPPING = {
    "West Bengal": "19-West Bengal",
    "Delhi": "07-Delhi",
    "Punjab": "03-Punjab",
    "Haryana": "06-Haryana",
    "Rajasthan": "08-Rajasthan",
    "Uttar Pradesh": "09-Uttar Pradesh",
    "Bihar": "10-Bihar",
    "Nagaland": "13-Nagaland",
    "Tripura": "16-Tripura",
    "Meghalaya": "17-Meghalaya",
    "Assam": "18-Assam",
    "Jharkhand": "20-Jharkhand",
    "Odisha": "21-Odisha",
    "Chhattisgarh": "22-Chhattisgarh",
    "Madhya Pradesh": "23-Madhya Pradesh",
    "Gujarat": "24-Gujarat",
    "Maharashtra": "27-Maharashtra",
    "Karnataka": "29-Karnataka",
    "Goa": "30-Goa",
    "Kerala": "32-Kerala",
    "Tamil Nadu": "33-Tamil Nadu",
    "Andaman & Nicobar Islands": "35-Andaman & Nicobar Islands",
    "Telangana": "36-Telangana",
    "Andhra Pradesh": "37-Andhra Pradesh"
}

# ============================================================
# HELPERS
# ============================================================

def load_template_from_github():
    r = requests.get(GITHUB_TEMPLATE_URL)
    if r.status_code != 200:
        st.error("❌ Template download failed")
        return None
    return io.BytesIO(r.content)

def process_file(file_data, data_type):
    df = pd.read_excel(file_data)
    df = df.rename(columns=COLUMN_MAPPING)
    df = df[list(COLUMN_MAPPING.values())]
    df["TYPE"] = data_type
    df["tcs_taxable_amount"] = pd.to_numeric(df["tcs_taxable_amount"], errors="coerce").fillna(0)
    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)

    if data_type == "Return":
        df["tcs_taxable_amount"] *= -1
        df["QTY"] *= -1

    return df

def calculate_tax_components(df, supplier_state_code):
    df["customer_state_code"] = df["J_mapped"].str[:2]
    intra = df["customer_state_code"] == supplier_state_code

    df["gst_rate"] = pd.to_numeric(df["gst_rate"], errors="coerce").fillna(0)
    tax = df["tcs_taxable_amount"] * df["gst_rate"] / 100

    df["CGST"] = tax.where(intra, 0) / 2
    df["SGST"] = tax.where(intra, 0) / 2
    df["IGST"] = tax.where(~intra, 0)

    return df

# ============================================================
# EXCEL GENERATION (UNCHANGED)
# ============================================================

def generate_combo_excel(df, template_stream):
    wb = load_workbook(template_stream)
    ws = wb["raw"]

    for r in range(3, ws.max_row + 1):
        for c in range(1, 16):
            ws.cell(r, c).value = None

    start_row = 3
    for i, row in enumerate(dataframe_to_rows(df[WRITE_COL_ORDER], index=False, header=False)):
        for j, v in enumerate(row):
            ws.cell(start_row + i, j + 2).value = v

        ws.cell(start_row + i, 1).value = "Messo"
        ws.cell(start_row + i, 10).value = df.loc[i, "J_mapped"]

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ============================================================
# CSV SUMMARIES (UNCHANGED)
# ============================================================

def generate_b2cs_csv(df):
    df = df.groupby(["J_mapped", "gst_rate"])["tcs_taxable_amount"].sum().reset_index()
    df["Type"] = "OE"
    df["Place Of Supply"] = df["J_mapped"]
    df["Rate"] = df["gst_rate"]
    df["Cess Amount"] = 0
    return df[["Type", "Place Of Supply", "Rate", "tcs_taxable_amount", "Cess Amount"]]\
        .rename(columns={"tcs_taxable_amount": "Taxable Value"})\
        .to_csv(index=False).encode()

def generate_hsn_summary(df):
    df = df.groupby(["hsn_code", "gst_rate"]).agg(
        Total_Quantity=("QTY", "sum"),
        Taxable_Value=("tcs_taxable_amount", "sum"),
        IGST=("IGST", "sum"),
        CGST=("CGST", "sum"),
        SGST=("SGST", "sum")
    ).reset_index()

    return df.to_csv(index=False).encode()

# ============================================================
# ✅ FIXED GST-OFFLINE JSON GENERATOR
# ============================================================

def generate_gstr1_json(df, gstin, fp):
    b2cs = []

    for (pos_name, rate), g in df.groupby(["J_mapped", "gst_rate"]):
        pos = pos_name[:2]

        txval = round(g["tcs_taxable_amount"].sum(), 2)
        iamt = round(g["IGST"].sum(), 2)
        camt = round(g["CGST"].sum(), 2)
        samt = round(g["SGST"].sum(), 2)

        if abs(txval) < 0.01:
            continue

        entry = {
            "sply_ty": "INTRA" if camt or samt else "INTER",
            "rt": int(rate),
            "typ": "OE",
            "pos": pos,
            "txval": txval,
            "csamt": 0
        }

        if camt or samt:
            entry["camt"] = camt
            entry["samt"] = samt
        else:
            entry["iamt"] = iamt

        b2cs.append(entry)

    hsn_b2c = []
    seq = 1

    for _, r in df.groupby(["hsn_code", "gst_rate"]).sum().reset_index().iterrows():
        hsn_b2c.append({
            "num": seq,
            "hsn_sc": str(int(r["hsn_code"])),
            "desc": "",
            "uqc": "NOS",
            "qty": round(r["QTY"], 3),
            "rt": int(r["gst_rate"]),
            "txval": round(r["tcs_taxable_amount"], 2),
            "iamt": round(r["IGST"], 2),
            "camt": round(r["CGST"], 2),
            "samt": round(r["SGST"], 2),
            "csamt": 0
        })
        seq += 1

    return json.dumps({
        "gstin": gstin,
        "fp": fp,
        "version": "GST3.2.3",
        "hash": "hash",
        "b2cs": b2cs,
        "hsn": {"hsn_b2c": hsn_b2c}
    }, indent=4).encode()
