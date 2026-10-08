import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="Rekapitulasi Absensi Siswa", layout="wide")

st.title("📋 Sistem Laporan Rekapitulasi Absensi Siswa")
st.caption("Aplikasi Rekapitulasi Absensi Otomatis Sesuai Format Standard Sekolah")

# Sidebar - Informasi Sekolah & Kelas
st.sidebar.header("🏫 Data Sekolah & Kelas")

uploaded_logo = st.sidebar.file_uploader("Upload Logo Sekolah (PNG / JPG)", type=["png", "jpg", "jpeg"], key="logo_uploader")
if uploaded_logo is not None:
    logo_img = Image.open(uploaded_logo)
    st.sidebar.image(logo_img, width=120, caption="Logo Sekolah")

nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SMP NEGERI 1 NANGA MAHAP")
tahun_pelajaran = st.sidebar.text_input("Tahun Pelajaran", "2025/2026")
kelas = st.sidebar.text_input("Kelas", "XI C")
bulan_tahun = st.sidebar.text_input("Bulan / Periode", "Oktober 2026")
tempat_cetak = st.sidebar.text_input("Kota / Tempat Laporan", "Nanga Mahap")
tgl_cetak = st.sidebar.date_input("Tanggal Cetak Laporan", datetime.today())
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "FRISKA EKASARI, S.Pd.")

bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
tgl_cetak_str = f"{tempat_cetak}, {tgl_cetak.day} {bulan_indo[tgl_cetak.month - 1]} {tgl_cetak.year}"

# Header Halaman Utama
st.markdown(f"<h2 style='text-align: center; margin-bottom: 0px;'>REKAPITULASI ABSENSI SISWA</h2>", unsafe_allow_html=True)
st.markdown(f"<h3 style='text-align: center; margin-top: 0px; margin-bottom: 0px;'>{nama_sekolah.upper()}</h3>", unsafe_allow_html=True)
st.markdown(f"<h4 style='text-align: center; margin-top: 0px;'>TAHUN PELAJARAN {tahun_pelajaran}</h4>", unsafe_allow_html=True)

st.markdown(f"**KELAS : {kelas.upper()}**")

# Data Default dari Gambar Sample
data_sample = [
    {"No": 1, "Nama Murid": "ABANG MUSHAWIR EDO", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 2, "Nama Murid": "AHMAD YANI", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 2},
    {"No": 3, "Nama Murid": "AL JAMI", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 1, "I": 0, "A": 3},
    {"No": 4, "Nama Murid": "AYU NINGSIH", "L/P": "P", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 5, "Nama Murid": "EPRI SASKIA", "L/P": "P", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 2},
    {"No": 6, "Nama Murid": "EVA DWI AGVENESA", "L/P": "P", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 1, "A": 0},
    {"No": 7, "Nama Murid": "FAKHRY LIANDRA WIJAYA", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 1, "I": 0, "A": 0},
    {"No": 8, "Nama Murid": "FATIRTA LAJESON", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 9, "Nama Murid": "FELISIA MONIK. S.L", "L/P": "P", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 1},
    {"No": 10, "Nama Murid": "FRANSISKUS EFRILDIO EVANO", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 11, "Nama Murid": "JIMI FAIZAL", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
]

if "data_absensi" not in st.session_state:
    st.session_state.data_absensi = pd.DataFrame(data_sample)

st.markdown("### 📝 Input / Edit Data Absensi")
edited_df = st.data_editor(
    st.session_state.data_absensi,
    column_config={
        "L/P": st.column_config.SelectboxColumn("L/P", options=["L", "P"], required=True),
        "HBE": st.column_config.NumberColumn("HBE", min_value=1, default=25),
        "S": st.column_config.NumberColumn("S (Sakit)", min_value=0, default=0),
        "I": st.column_config.NumberColumn("I (Izin)", min_value=0, default=0),
        "A": st.column_config.NumberColumn("A (Alpa)", min_value=0, default=0),
    },
    num_rows="dynamic",
    use_container_width=True,
    key="editor_absensi"
)

# Kalkulasi Data
df_calc = edited_df.copy()
for col in ["HBE", "S", "I", "A"]:
    df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)

df_calc["JUMLAH"] = df_calc["S"] + df_calc["I"] + df_calc["A"]
df_calc["JUMLAH HADIR"] = df_calc["HBE"] - df_calc["JUMLAH"]
df_calc["JUMLAH HADIR"] = df_calc["JUMLAH HADIR"].apply(lambda x: max(0, x))

df_calc["PRESENTASE S"] = (df_calc["S"] / df_calc["HBE"])
df_calc["PRESENTASE I"] = (df_calc["I"] / df_calc["HBE"])
df_calc["PRESENTASE A"] = (df_calc["A"] / df_calc["HBE"])
df_calc["PRESENTASE KEHADIRAN %"] = (df_calc["JUMLAH HADIR"] / df_calc["HBE"])

# Tampilan Web Sesuai Format Tabel Foto
df_view = pd.DataFrame()
df_view["NO"] = df_calc["No"]
df_view["NAMA MURID"] = df_calc["Nama Murid"]
df_view["L/P"] = df_calc["L/P"]
df_view["NOMOR INDUK"] = df_calc["Nomor Induk"]
df_view["HBE"] = df_calc["HBE"]
df_view["ABSENSI (S)"] = df_calc["S"]
df_view["ABSENSI (I)"] = df_calc["I"]
df_view["ABSENSI (A)"] = df_calc["A"]
df_view["JUMLAH"] = df_calc["JUMLAH"]
df_view["JUMLAH HADIR"] = df_calc["JUMLAH HADIR"]
df_view["PRESENTASE (S)"] = (df_calc["PRESENTASE S"] * 100).round(0).astype(int).astype(str) + "%"
df_view["PRESENTASE (I)"] = (df_calc["PRESENTASE I"] * 100).round(0).astype(int).astype(str) + "%"
df_view["PRESENTASE (A)"] = (df_calc["PRESENTASE A"] * 100).round(0).astype(int).astype(str) + "%"
df_view["PRESENTASE KEHADIRAN %"] = (df_calc["PRESENTASE KEHADIRAN %"] * 100).round(0).astype(int).astype(str) + "%"

st.markdown("### 📊 Hasil Rekapitulasi Absensi Siswa")
st.dataframe(df_view, use_container_width=True)

# Function Generate Excel Sesuai Format Foto Laporan
def generate_excel_foto_format():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rekap Absensi"

    # Styling Font & Color
    font_title = Font(name="Calibri", size=14, bold=True)
    font_subtitle = Font(name="Calibri", size=11, bold=True)
    font_header = Font(name="Calibri", size=10, bold=True)
    font_data = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid") # Warna Hijau L/P
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    # Title Block
    ws.merge_cells("A1:N1")
    ws.cell(row=1, column=1, value="REKAPITULASI ABSENSI SISWA").font = font_title
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:N2")
    ws.cell(row=2, column=1, value=nama_sekolah.upper()).font = font_title
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A3:N3")
    ws.cell(row=3, column=1, value=f"TAHUN PELAJARAN {tahun_pelajaran}").font = font_subtitle
    ws.cell(row=3, column=1).alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=5, column=1, value=f"KELAS   : {kelas.upper()}").font = font_bold

    # Header Row 7 & 8 (Multi-header bertingkat)
    headers_r7 = [
        ("NO", "A7", "A8"),
        ("NAMA MURID", "B7", "B8"),
        ("L/P", "C7", "C8"),
        ("NOMOR INDUK", "D7", "D8"),
        ("HBE", "E7", "E8"),
        ("ABSENSI", "F7", "H7"),
        ("JUMLAH", "I7", "I8"),
        ("JUMLAH HADIR", "J7", "J8"),
        ("PRESENTASE", "K7", "M7"),
        ("PRESENTASE KEHADIRAN %", "N7", "N8")
    ]

    for title, start_col, end_col in headers_r7:
        if start_col != end_col:
            ws.merge_cells(f"{start_col}:{end_col}")
        cell = ws[start_col.split(":")[0]]
        cell.value = title
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Sub-headers Row 8
    ws.cell(row=8, column=6, value="S").font = font_header
    ws.cell(row=8, column=7, value="I").font = font_header
    ws.cell(row=8, column=8, value="A").font = font_header

    ws.cell(row=8, column=11, value="S").font = font_header
    ws.cell(row=8, column=12, value="I").font = font_header
    ws.cell(row=8, column=13, value="A").font = font_header

    for c in range(1, 15):
        cell_r7 = ws.cell(row=7, column=c)
        cell_r8 = ws.cell(row=8, column=c)
        cell_r7.border = thin_border
        cell_r8.border = thin_border
        cell_r8.alignment = Alignment(horizontal="center", vertical="center")

    # Data Rows (Mulai Baris 9)
    start_row = 9
    for idx, row in df_calc.iterrows():
        r = start_row + idx
        ws.cell(row=r, column=1, value=row.get("No", idx + 1))
        ws.cell(row=r, column=2, value=row.get("Nama Murid", ""))
        
        lp_cell = ws.cell(row=r, column=3, value=row.get("L/P", ""))
        lp_cell.fill = green_fill # Sorot Warna Hijau L/P
        
        ws.cell(row=r, column=4, value=row.get("Nomor Induk", ""))
        ws.cell(row=r, column=5, value=row.get("HBE", 25))
        ws.cell(row=r, column=6, value=row.get("S", 0))
        ws.cell(row=r, column=7, value=row.get("I", 0))
        ws.cell(row=r, column=8, value=row.get("A", 0))
        
        # Formula Excel Otomatis
        ws.cell(row=r, column=9, value=f"=SUM(F{r}:H{r})")
        ws.cell(row=r, column=10, value=f"=E{r}-I{r}")
        
        ws.cell(row=r, column=11, value=f"=F{r}/E{r}").number_format = '0%'
        ws.cell(row=r, column=12, value=f"=G{r}/E{r}").number_format = '0%'
        ws.cell(row=r, column=13, value=f"=H{r}/E{r}").number_format = '0%'
        ws.cell(row=r, column=14, value=f"=J{r}/E{r}").number_format = '0%'

        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.font = font_bold if c in [1, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14] else font_data
            
            if c in [2, 4]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Autofit Width
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row >= 7:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 6)

    ws.column_dimensions['B'].width = 30 # Lebar Nama Murid
    ws.column_dimensions['D'].width = 18 # Lebar Nomor Induk

    # Pengaturan Siap Cetak (Print Ready)
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    wb.save(output)
    return output.getvalue()

st.download_button(
    label="📥 Download File Excel Rekap (Persis Format Foto & Siap Cetak)",
    data=generate_excel_foto_format(),
    file_name=f"Rekap_Absensi_Siswa_{kelas}_{tahun_pelajaran.replace('/', '-')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
