import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
from datetime import datetime
from PIL import Image

st.set_page_config(page_title="Sistem Absensi & Mutasi Siswa SMP", layout="wide")

st.title("📋 Sistem Laporan Absensi & Mutasi Siswa SMP")
st.caption("Aplikasi Rekapitulasi Absensi Bulanan, Harian, dan Mutasi Siswa Otomatis")

# Sidebar - Informasi Sekolah & Kelas
st.sidebar.header("🏫 Data Sekolah & Kelas (SMP)")

# --- FITUR UPLOAD LOGO SEKOLAH ---
uploaded_logo = st.sidebar.file_uploader("Upload Logo Sekolah (PNG / JPG)", type=["png", "jpg", "jpeg"], key="logo_uploader")

if uploaded_logo is not None:
    logo_img = Image.open(uploaded_logo)
    st.sidebar.image(logo_img, width=120, caption="Logo Sekolah")

nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SMP NEGERI 1 NANGA MAHAP")

# --- FITUR PILIHAN KELAS JENJANG SMP ---
opsi_tingkat = ["Kelas 7", "Kelas 8", "Kelas 9", "Lainnya (Ketik Manual)"]
pilihan_tingkat = st.sidebar.selectbox("Pilih Tingkat Kelas", opsi_tingkat, index=0)

if pilihan_tingkat == "Lainnya (Ketik Manual)":
    kelas = st.sidebar.text_input("Ketik Nama Kelas Kustom", "Kelas 7A")
else:
    rombel = st.sidebar.selectbox("Pilih Rombel / Abjad Kelas", ["A", "B", "C", "D", "E", "F", "G", "H", "Tanpa Rombel"], index=0)
    if rombel == "Tanpa Rombel":
        kelas = pilihan_tingkat
    else:
        kelas = f"{pilihan_tingkat} {rombel}"

bulan_tahun = st.sidebar.text_input("Bulan / Periode", "Oktober 2026")
tempat_cetak = st.sidebar.text_input("Kota / Tempat Laporan", "Nanga Mahap")
tgl_cetak = st.sidebar.date_input("Tanggal Cetak Laporan", datetime.today())
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "FRISKA EKASARI, S.Pd.")

# Helper Function: Formatting Tanggal Indonesia
bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
tgl_cetak_str = f"{tempat_cetak}, {tgl_cetak.day} {bulan_indo[tgl_cetak.month - 1]} {tgl_cetak.year}"

# Helper Function: Styling Excel, Logo & TTD
def apply_excel_styling(ws, headers, title_1, title_2, title_3, wali_kelas_nama, tgl_str, logo_file=None, start_data_row=6):
    title_font = Font(name="Arial", size=11, bold=True)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    if logo_file is not None:
        try:
            logo_file.seek(0)
            img = OpenpyxlImage(logo_file)
            img.width = 65
            img.height = 65
            ws.add_image(img, "A1")
            ws.row_dimensions[1].height = 20
            ws.row_dimensions[2].height = 20
            ws.row_dimensions[3].height = 20
        except Exception:
            pass

    ws.cell(row=1, column=2, value=title_1).font = title_font
    ws.cell(row=2, column=2, value=title_2).font = title_font
    ws.cell(row=3, column=2, value=title_3).font = title_font

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=start_data_row - 1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    last_row = ws.max_row
    for r in range(start_data_row - 1, last_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if r >= start_data_row and c in [1, 2, 4, 5, 6, 7, 8, 9, 10, 11]:
                cell.alignment = Alignment(horizontal="center")

    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=last_row, column=c)
        if "JUMLAH" in str(ws.cell(row=last_row, column=1).value or "").upper():
            cell.font = Font(name="Arial", size=10, bold=True)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ws.append([])
    ws.append([])
    ttd_col = max(1, len(headers) - 3)
    ws.cell(row=ws.max_row + 1, column=ttd_col, value=tgl_str)
    ws.cell(row=ws.max_row + 1, column=ttd_col, value="Wali Kelas,")
    ws.append([])
    ws.append([])
    ws.cell(row=ws.max_row + 1, column=ttd_col, value=f"({wali_kelas_nama})").font = Font(bold=True)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Header Halaman Utama dengan Logo
col_logo, col_header = st.columns([1, 6])
with col_logo:
    if uploaded_logo is not None:
        st.image(uploaded_logo, width=90)
with col_header:
    st.markdown(f"### **{nama_sekolah.upper()}**")
    st.markdown(f"**{kelas.upper()}** | Periode: **{bulan_tahun}**")

# Tab Fitur
tab1, tab2, tab3 = st.tabs([
    "📊 1. Rekapitulasi Absensi Bulanan", 
    "📅 2. Kehadiran Harian Siswa", 
    "🔄 3. Mutasi Siswa"
])

# ==========================================
# TAB 1: REKAPITULASI ABSENSI BULANAN
# ==========================================
with tab1:
    st.subheader("📊 Rekapitulasi Absensi Bulanan Siswa SMP")
    
    st.markdown("### 📤 Upload Data Siswa dari Excel / CSV")
    uploaded_file = st.file_uploader(
        "Pilih file Excel (.xlsx) atau CSV (.csv) berisi daftar siswa", 
        type=["xlsx", "csv"],
        key="uploader_bulanan"
    )

    if "data_bulanan" not in st.session_state:
        st.session_state.data_bulanan = pd.DataFrame({
            "No": [1, 2, 3],
            "Nomor Induk / NISN": ["1001", "1002", "1003"],
            "Nama Siswa": ["Ahmad Fauzi", "Budi Santoso", "Citra Dewi"],
            "Jenis Kelamin": ["L", "L", "P"],
            "HBE": [24, 24, 24],
            "Sakit": [1, 0, 0],
            "Izin": [0, 1, 0],
            "Alpa": [0, 0, 1]
        })

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df_upload = pd.read_csv(uploaded_file)
            else:
                df_upload = pd.read_excel(uploaded_file)
            
            required_cols = ["No", "Nomor Induk / NISN", "Nama Siswa", "Jenis Kelamin", "HBE", "Sakit", "Izin", "Alpa"]
            for col in required_cols:
                if col not in df_upload.columns:
                    if col in ["Sakit", "Izin", "Alpa"]:
                        df_upload[col] = 0
                    elif col == "HBE":
                        df_upload[col] = 24
                    elif col == "Jenis Kelamin":
                        df_upload[col] = "L"
            
            st.session_state.data_bulanan = df_upload[required_cols]
            st.success("✅ Data siswa berhasil diunggah!")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    st.write("Silakan sesuaikan data kehadiran pada tabel interaktif di bawah ini:")

    edited_df = st.data_editor(
        st.session_state.data_bulanan,
        column_config={
            "Jenis Kelamin": st.column_config.SelectboxColumn("Jenis Kelamin", options=["L", "P"], required=True),
            "HBE": st.column_config.NumberColumn("HBE (Hari Belajar Efektif)", min_value=1, default=24),
            "Sakit": st.column_config.NumberColumn("Sakit", min_value=0, default=0),
            "Izin": st.column_config.NumberColumn("Izin", min_value=0, default=0),
            "Alpa": st.column_config.NumberColumn("Alpa", min_value=0, default=0),
        },
        num_rows="dynamic",
        use_container_width=True,
        key="editor_bulanan"
    )

    df_calc = edited_df.copy()
    for col in ["HBE", "Sakit", "Izin", "Alpa"]:
        df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)

    df_calc["Jumlah Ketidakhadiran"] = df_calc["Sakit"] + df_calc["Izin"] + df_calc["Alpa"]
    df_calc["Jumlah Hadir"] = df_calc["HBE"] - df_calc["Jumlah Ketidakhadiran"]
    df_calc["Jumlah Hadir"] = df_calc["Jumlah Hadir"].apply(lambda x: max(0, x))
    
    df_calc["% Kehadiran"] = (df_calc["Jumlah Hadir"] / df_calc["HBE"])
    df_calc_display = df_calc.copy()
    df_calc_display["% Kehadiran"] = (df_calc_display["% Kehadiran"] * 100).round(1).apply(lambda x: f"{x}%")

    st.markdown("### 📈 Tabel Hasil Rekapitulasi Otomatis")
    st.dataframe(df_calc_display, use_container_width=True)

    def generate_excel_rekap_bulanan():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Absensi Bulanan"

        headers = ["No", "Nomor Induk", "Nama Siswa", "JK", "HBE", "Sakit", "Izin", "Alpa", "Jumlah Absen", "Jumlah Hadir", "% Kehadiran"]
        
        ws.append([nama_sekolah.upper()])
        ws.append([f"LAPORAN REKAPITULASI ABSENSI SISWA - {kelas.upper()}"])
        ws.append([f"PERIODE: {bulan_tahun.upper()}"])
        ws.append([])
        ws.append(headers)

        start_data_row = 6
        for idx, row in df_calc.iterrows():
            curr_row = start_data_row + idx
            ws.append([
                row.get("No", ""),
                row.get("Nomor Induk / NISN", ""),
                row.get("Nama Siswa", ""),
                row.get("Jenis Kelamin", ""),
                row.get("HBE", 24),
                row.get("Sakit", 0),
