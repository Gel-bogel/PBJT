import streamlit as st
import pandas as pd
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Konfigurasi Halaman Streamlit
st.set_page_config(page_title="Laporan Realisasi PPJ / PBJT-TL", layout="wide")

st.title("⚡ Laporan Realisasi Hasil Pemungutan PPJ / PBJT-TL")
st.subheader("PT PLN (PERSERO) UP3 MEDAN UTARA")
st.caption("Konsolidasi 5 ULP (Format Baru AP2T)")

# ----------------------------------------------------
# FUNGSI EKSTRAKSI DATA FORMAT BARU AP2T
# ----------------------------------------------------

def extract_ppj1_pemda_new(df1, kota):
    """
    Ekstraksi data Halaman 1 (PPJ 1) Format Baru AP2T.
    Target Kolom PBJT:
    - k3 (Lunas Online / DPH22)  : Indeks kolom 17
    - k4 (Lunas Offline / DPH21) : Indeks kolom 13
    - k5 (Beban Kantor / DPH24)  : Indeks kolom 25
    - k6 (Lunas Terpusat / DPH25): Indeks kolom 29
    """
    default_res = {"k3": 0, "k4": 0, "k5": 0, "k6": 0}
    if df1 is None or df1.empty:
        return default_res

    mask = df1.iloc[:, 1].astype(str).str.contains(kota, case=False, na=False)
    df_pemda = df1[mask]

    if df_pemda.empty:
        return default_res

    k4_val = pd.to_numeric(df_pemda.iloc[:, 13], errors='coerce').sum() if df1.shape[1] > 13 else 0
    k3_val = pd.to_numeric(df_pemda.iloc[:, 17], errors='coerce').sum() if df1.shape[1] > 17 else 0
    k5_val = pd.to_numeric(df_pemda.iloc[:, 25], errors='coerce').sum() if df1.shape[1] > 25 else 0
    k6_val = pd.to_numeric(df_pemda.iloc[:, 29], errors='coerce').sum() if df1.shape[1] > 29 else 0

    return {
        "k3": k3_val,
        "k4": k4_val,
        "k5": k5_val,
        "k6": k6_val
    }


def extract_ppj2_pemda_new(df2, kota):
    """
    Ekstraksi data Halaman 2 (PPJ 2) Format Baru AP2T.
    Target Kolom PBJT:
    - k7  (Restitusi Taglis)          : Indeks kolom 12
    - k8  (Non Taglis Online)         : Indeks kolom 24
    - k10 (Non Taglis Offline / I-06) : Indeks kolom 28
    - k11 (Nota Buku)                 : Indeks kolom 32
    - k9  (Prabayar / Token)          : Indeks kolom 36
    - k14 (Kompensasi TMP)            : Indeks kolom 40
    - k15 (Kompensasi UJL)            : Indeks kolom 44
    - k13 (Total Pelunasan)           : Indeks kolom 47
    """
    default_res = {
        "k7": 0, "k8": 0, "k9": 0, "k10": 0, "k11": 0, "k12": 0, "k13": 0, "k14": 0, "k15": 0
    }
    if df2 is None or df2.empty:
        return default_res

    mask = df2.iloc[:, 1].astype(str).str.contains(kota, case=False, na=False)
    df_pemda = df2[mask]

    if df_pemda.empty:
        return default_res

    k7_val  = pd.to_numeric(df_pemda.iloc[:, 12], errors='coerce').sum() if df2.shape[1] > 12 else 0
    k8_val  = pd.to_numeric(df_pemda.iloc[:, 24], errors='coerce').sum() if df2.shape[1] > 24 else 0
    k10_val = pd.to_numeric(df_pemda.iloc[:, 28], errors='coerce').sum() if df2.shape[1] > 28 else 0
    k11_val = pd.to_numeric(df_pemda.iloc[:, 32], errors='coerce').sum() if df2.shape[1] > 32 else 0
    k9_val  = pd.to_numeric(df_pemda.iloc[:, 36], errors='coerce').sum() if df2.shape[1] > 36 else 0
    k14_val = pd.to_numeric(df_pemda.iloc[:, 40], errors='coerce').sum() if df2.shape[1] > 40 else 0
    k15_val = pd.to_numeric(df_pemda.iloc[:, 44], errors='coerce').sum() if df2.shape[1] > 44 else 0

    target_tot_idx = 47 if df2.shape[1] > 47 else df2.shape[1] - 1
    k13_val = pd.to_numeric(df_pemda.iloc[:, target_tot_idx], errors='coerce').sum()

    return {
        "k7": k7_val,
        "k8": k8_val,
        "k9": k9_val,
        "k10": k10_val,
        "k11": k11_val,
        "k12": 0,
        "k13": k13_val,
        "k14": k14_val,
        "k15": k15_val
    }


# ----------------------------------------------------
# SIDEBAR / UNGGAH BERKAS
# ----------------------------------------------------
st.sidebar.header("📁 Unggah Berkas AP2T (Format Baru)")

list_ulp = ["ULP BELAWAN", "ULP DENAI", "ULP HELVETIA", "ULP LABUHAN", "ULP MEDAN TIMUR"]

uploaded_files_hal1 = st.sidebar.file_uploader(
    "1. Berkas Halaman 1 (PPJ 1)", type=["xlsx", "xls"], accept_multiple_files=True
)
uploaded_files_hal2 = st.sidebar.file_uploader(
    "2. Berkas Halaman 2 (PPJ 2)", type=["xlsx", "xls"], accept_multiple_files=True
)


def find_file_by_ulp(file_list, ulp_name):
    if not file_list:
        return None
    clean_ulp = ulp_name.replace("ULP ", "").strip()
    for f in file_list:
        if clean_ulp.lower() in f.name.lower():
            try:
                return pd.read_excel(f, header=None)
            except Exception:
                return None
    return None


# ----------------------------------------------------
# PEMROSESAN DATA
# ----------------------------------------------------
data_total = []
data_medan = []
data_deliserdang = []

for idx, ulp in enumerate(list_ulp, 1):
    df1 = find_file_by_ulp(uploaded_files_hal1, ulp)
    df2 = find_file_by_ulp(uploaded_files_hal2, ulp)

    row_m = {"NO": idx, "NAMA ULP": ulp}
    row_ds = {"NO": idx, "NAMA ULP": ulp}

    d1_m = extract_ppj1_pemda_new(df1, "KOTA MEDAN") if df1 is not None else {"k3": 0, "k4": 0, "k5": 0, "k6": 0}
    d1_ds = extract_ppj1_pemda_new(df1, "DELI SERDANG") if df1 is not None else {"k3": 0, "k4": 0, "k5": 0, "k6": 0}

    d2_m = extract_ppj2_pemda_new(df2, "KOTA MEDAN") if df2 is not None else {"k7": 0, "k8": 0, "k9": 0, "k10": 0, "k11": 0, "k12": 0, "k13": 0, "k14": 0, "k15": 0}
    d2_ds = extract_ppj2_pemda_new(df2, "DELI SERDANG") if df2 is not None else {"k7": 0, "k8": 0, "k9": 0, "k10": 0, "k11": 0, "k12": 0, "k13": 0, "k14": 0, "k15": 0}

    row_m.update(d1_m)
    row_m.update(d2_m)

    row_ds.update(d1_ds)
    row_ds.update(d2_ds)

    row_tot = {"NO": idx, "NAMA ULP": ulp}
    for key in ["k3", "k4", "k5", "k6", "k7", "k8", "k9", "k10", "k11", "k12", "k13", "k14", "k15"]:
        row_tot[key] = row_m.get(key, 0) + row_ds.get(key, 0)

    data_total.append(row_tot)
    data_medan.append(row_m)
    data_deliserdang.append(row_ds)

df_tot = pd.DataFrame(data_total)
df_m = pd.DataFrame(data_medan)
df_ds = pd.DataFrame(data_deliserdang)


# ----------------------------------------------------
# FUNGSI RENDER TABEL HTML LENGKAP KUSTOM
# ----------------------------------------------------
def render_custom_html_table(df):
    if df.empty:
        return ""

    headers = [
        ("1", "NO"),
        ("2", "NAMA ULP"),
        ("3", "LUNAS ONLINE (KODE GERAK 22)"),
        ("4", "LUNAS OFFLINE (KODE GERAK 21)"),
        ("5", "BEBAN KANTOR (KODE GERAK 24)"),
        ("6", "LUNAS TERPUSAT (KODE GERAK 25)"),
        ("7", "RESTITUSI"),
        ("8", "NON TAGLIS ONLINE"),
        ("9", "NON TAGLIS OFFLINE"),
        ("10", "NOTA BUKU"),
        ("11", "PRABAYAR"),
        ("12", "KOMPENSASI TMP"),
        ("13", "KOMPENSASI UJL"),
        ("14", "TOTAL PELUNASAN")
    ]

    keys = ["k3", "k4", "k5", "k6", "k7", "k8", "k10", "k11", "k9", "k14", "k15", "k13"]

    html = """
    <style>
        .custom-table-container {
            width: 100%;
            overflow-x: auto;
            margin-bottom: 25px;
        }
        .custom-table {
            width: 100%;
            border-collapse: collapse;
            font-family: sans-serif;
            font-size: 13px;
            background-color: #FFFFFF;
            color: #000000;
        }
        .custom-table th {
            background-color: #9CD2EC;
            color: #000000;
            border: 1px solid #73A9C2;
            padding: 8px 5px;
            text-align: center;
            font-weight: bold;
            white-space: nowrap;
        }
        .custom-table td {
            border: 1px solid #B0C4DE;
            padding: 6px 8px;
            vertical-align: middle;
            white-space: nowrap;
        }
        .num-col {
            display: flex;
            justify-content: space-between;
            min-width: 110px;
        }
        .total-row {
            background-color: #62B8DF !important;
            font-weight: bold;
        }
    </style>
    <div class="custom-table-container">
    <table class="custom-table">
        <thead>
            <tr>
    """

    # Baris Nomor Kolom (1, 2, 3, dst.)
    for num, _ in headers:
        html += f"<th>{num}</th>"
    html += "</tr><tr>"

    # Baris Nama Header
    for _, title in headers:
        html += f"<th>{title}</th>"
    html += "</tr></thead><tbody>"

    # Isian Data Baris ULP
    totals = {k: df[k].sum() for k in keys if k in df.columns}

    for _, row in df.iterrows():
        html += f"<tr><td style='text-align:center;'>{row['NO']}</td><td>{row['NAMA ULP']}</td>"
        for k in keys:
            val = row.get(k, 0)
            formatted_val = f"{int(val):,}".replace(",", ".")
            html += f"<td><div class='num-col'><span>Rp</span><span>{formatted_val}</span></div></td>"
        html += "</tr>"

    # Baris Jumlah Total UP3 Medan Utara
    html += f"<tr class='total-row'><td style='text-align:center;'></td><td>JUMLAH TOTAL UP3 MEDAN UTARA</td>"
    for k in keys:
        tot_val = totals.get(k, 0)
        formatted_tot = f"{int(tot_val):,}".replace(",", ".")
        html += f"<td><div class='num-col'><span>Rp</span><span>{formatted_tot}</span></div></td>"
    html += "</tr>"

    html += "</tbody></table></div>"
    return html


# ----------------------------------------------------
# TAMPILAN DASHBOARD STREAMLIT
# ----------------------------------------------------
st.subheader("🏢 TABEL 1. TOTAL KESELURUHAN (PEMDA MEDAN + DELI SERDANG)")
st.markdown(render_custom_html_table(df_tot), unsafe_allow_html=True)

st.subheader("🏛️ TABEL 2. PEMDA KOTA MEDAN")
st.markdown(render_custom_html_table(df_m), unsafe_allow_html=True)

st.subheader("🌾 TABEL 3. PEMDA KABUPATEN DELI SERDANG")
st.markdown(render_custom_html_table(df_ds), unsafe_allow_html=True)


# ----------------------------------------------------
# EXPORT KE EXCEL DENGAN OPENPYXL STYLING
# ----------------------------------------------------
def export_to_excel(df_tot, df_m, df_ds):
    output = io.BytesIO()
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Konsolidasi Total"
    ws2 = wb.create_sheet(title="Kota Medan")
    ws3 = wb.create_sheet(title="Deli Serdang")

    sheets = [(ws1, df_tot), (ws2, df_m), (ws3, df_ds)]

    headers = [
        "NO", "NAMA ULP", "LUNAS ONLINE", "LUNAS OFFLINE", "BEBAN KANTOR", 
        "LUNAS TERPUSAT", "RESTITUSI", "NON TAGLIS ONLINE", "NON TAGLIS OFFLINE", 
        "NOTA BUKU", "PRABAYAR", "KOMPENSASI TMP", "KOMPENSASI UJL", "TOTAL PELUNASAN"
    ]
    keys = ["k3", "k4", "k5", "k6", "k7", "k8", "k10", "k11", "k9", "k14", "k15", "k13"]

    for ws, df_curr in sheets:
        ws.append(headers)
        
        for _, r in df_curr.iterrows():
            ws.append([r["NO"], r["NAMA ULP"]] + [r.get(k, 0) for k in keys])
            
        tot_row = ["", "JUMLAH TOTAL UP3 MEDAN UTARA"] + [df_curr[k].sum() for k in keys]
        ws.append(tot_row)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=3, max_col=len(headers)):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0'

    wb.save(output)
    output.seek(0)
    return output


st.markdown("---")
excel_file = export_to_excel(df_tot, df_m, df_ds)

st.download_button(
    label="📥 Unduh Laporan Consolidated Excel (.xlsx)",
    data=excel_file,
    file_name="Laporan_Realisasi_PPJ_PBJT_UP3_Medan_Utara.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
