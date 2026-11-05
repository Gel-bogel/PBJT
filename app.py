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
    default_res = {"k7": 0, "k8": 0, "k9": 0, "k10": 0, "k11": 0, "k12": 0, "k13": 0}
    if df2 is None or df2.empty:
        return default_res

    mask = df2.iloc[:, 1].astype(str).str.contains(kota, case=False, na=False)
    df_pemda = df2[mask]

    if df_pemda.empty:
        return default_res

    k7_val = pd.to_numeric(df_pemda.iloc[:, 12], errors='coerce').sum() if df2.shape[1] > 12 else 0
    k8_val = pd.to_numeric(df_pemda.iloc[:, 24], errors='coerce').sum() if df2.shape[1] > 24 else 0
    k9_val = pd.to_numeric(df_pemda.iloc[:, 36], errors='coerce').sum() if df2.shape[1] > 36 else 0

    target_tot_idx = 47 if df2.shape[1] > 47 else df2.shape[1] - 1
    k13_val = pd.to_numeric(df_pemda.iloc[:, target_tot_idx], errors='coerce').sum()

    return {
        "k7": k7_val,
        "k8": k8_val,
        "k9": k9_val,
        "k10": 0,
        "k11": 0,
        "k12": 0,
        "k13": k13_val
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

    row_m = {"NO": str(idx), "NAMA ULP": ulp}
    row_ds = {"NO": str(idx), "NAMA ULP": ulp}

    d1_m = extract_ppj1_pemda_new(df1, "KOTA MEDAN") if df1 is not None else {"k3": 0, "k4": 0, "k5": 0, "k6": 0}
    d1_ds = extract_ppj1_pemda_new(df1, "DELI SERDANG") if df1 is not None else {"k3": 0, "k4": 0, "k5": 0, "k6": 0}

    d2_m = extract_ppj2_pemda_new(df2, "KOTA MEDAN") if df2 is not None else {"k7": 0, "k8": 0, "k9": 0, "k10": 0, "k11": 0, "k12": 0, "k13": 0}
    d2_ds = extract_ppj2_pemda_new(df2, "DELI SERDANG") if df2 is not None else {"k7": 0, "k8": 0, "k9": 0, "k10": 0, "k11": 0, "k12": 0, "k13": 0}

    row_m.update(d1_m)
    row_m.update(d2_m)

    row_ds.update(d1_ds)
    row_ds.update(d2_ds)

    row_tot = {"NO": str(idx), "NAMA ULP": ulp}
    for key in ["k3", "k4", "k5", "k6", "k7", "k8", "k9", "k10", "k11", "k12", "k13"]:
        row_tot[key] = row_m.get(key, 0) + row_ds.get(key, 0)

    data_total.append(row_tot)
    data_medan.append(row_m)
    data_deliserdang.append(row_ds)

df_tot = pd.DataFrame(data_total)
df_m = pd.DataFrame(data_medan)
df_ds = pd.DataFrame(data_deliserdang)


def format_dataframe_display(df):
    if df.empty:
        return df

    # Hitung jumlah total sebelum dikonversi ke string
    tot_row = {
        "NO": "",
        "NAMA ULP": "JUMLAH TOTAL UP3 MEDAN UTARA",
    }
    for col in ["k3", "k4", "k5", "k6", "k7", "k8", "k9", "k10", "k11", "k12", "k13"]:
        tot_row[col] = df[col].sum() if col in df.columns else 0

    df_out = pd.concat([df, pd.DataFrame([tot_row])], ignore_index=True)

    # Format angka menjadi "Rp 1.000.000" agar RATA KIRI
    for col in ["k3", "k4", "k5", "k6", "k7", "k8", "k9", "k10", "k11", "k12", "k13"]:
        if col in df_out.columns:
            df_out[col] = df_out[col].apply(lambda x: f"Rp {int(x):,}".replace(",", "."))

    df_out = df_out.rename(columns={
        "k3": "LUNAS ONLINE (KODE GERAK 22)",
        "k4": "LUNAS OFFLINE (KODE GERAK 21)",
        "k5": "BEBAN KANTOR (KODE GERAK 24)",
        "k6": "LUNAS TERPUSAT (KODE GERAK 25)",
        "k7": "RESTITUSI",
        "k8": "NON TAGLIS ONLINE",
        "k9": "PRABAYAR",
        "k10": "NON TAGLIS OFFLINE",
        "k11": "NOTA BUKU",
        "k12": "LAIN-LAIN",
        "k13": "TOTAL PELUNASAN"
    })
    return df_out


# Fungsi Styling Warna Header Biru & Kontras Baris Total
def apply_custom_table_style(df):
    num_cols = [c for c in df.columns if c not in ["NO", "NAMA ULP"]]
    
    styler = df.style.set_table_styles([
        {
            'selector': 'th',
            'props': [
                ('background-color', '#1F4E79'),  # Header Biru Tua
                ('color', '#FFFFFF'),              # Teks Putih
                ('font-weight', 'bold'),
                ('text-align', 'center'),
                ('border', '1px solid #102A43')
            ]
        },
        {
            'selector': 'td',
            'props': [
                ('text-align', 'left'),            # Teks & Rp Rata Kiri
                ('padding', '8px')
            ]
        }
    ])

    # Highlight baris Jumlah Total dengan Warna Kontras Biru Cerah/Kuning Emas
    def highlight_total(row):
        if row['NAMA ULP'] == 'JUMLAH TOTAL UP3 MEDAN UTARA':
            return ['background-color: #2E5B88; color: #FFD700; font-weight: bold;'] * len(row)
        return [''] * len(row)

    styler = styler.apply(highlight_total, axis=1)
    return styler


# ----------------------------------------------------
# TAMPILAN TABEL DASHBOARD STREAMLIT
# ----------------------------------------------------
df_tot_display = format_dataframe_display(df_tot)
df_m_display = format_dataframe_display(df_m)
df_ds_display = format_dataframe_display(df_ds)

st.subheader("🏢 TABEL 1. TOTAL KESELURUHAN (PEMDA MEDAN + DELI SERDANG)")
st.dataframe(apply_custom_table_style(df_tot_display), use_container_width=True, hide_index=True)

st.subheader("🏛️ TABEL 2. PEMDA KOTA MEDAN")
st.dataframe(apply_custom_table_style(df_m_display), use_container_width=True, hide_index=True)

st.subheader("🌾 TABEL 3. PEMDA KABUPATEN DELI SERDANG")
st.dataframe(apply_custom_table_style(df_ds_display), use_container_width=True, hide_index=True)


# ----------------------------------------------------
# EXPORT KE EXCEL DENGAN OPENPYXL STYLING
# ----------------------------------------------------
def export_to_excel(df_total_disp, df_medan_disp, df_ds_disp):
    output = io.BytesIO()
    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Konsolidasi Total"
    ws2 = wb.create_sheet(title="Kota Medan")
    ws3 = wb.create_sheet(title="Deli Serdang")

    sheets_data = [
        (ws1, "LAPORAN KONSOLIDASI TOTAL UP3 MEDAN UTARA", df_total_disp),
        (ws2, "LAPORAN REALISASI PEMDA KOTA MEDAN", df_medan_disp),
        (ws3, "LAPORAN REALISASI PEMDA KABUPATEN DELI SERDANG", df_ds_disp),
    ]

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    for ws, title, df_data in sheets_data:
        ws.append([title])
        ws.append([])

        headers = list(df_data.columns)
        ws.append(headers)

        header_row_idx = 3
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=header_row_idx, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row in df_data.itertuples(index=False):
            ws.append(list(row))

        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=len(headers)):
            for cell in row:
                cell.border = thin_border

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    wb.save(output)
    output.seek(0)
    return output


st.markdown("---")
excel_file = export_to_excel(df_tot_display, df_m_display, df_ds_display)

st.download_button(
    label="📥 Unduh Laporan Consolidated Excel (.xlsx)",
    data=excel_file,
    file_name="Laporan_Realisasi_PPJ_PBJT_UP3_Medan_Utara.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
