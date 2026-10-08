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
                try:
                    df_upload = pd.read_csv(uploaded_file, sep=None, engine='python')
                except Exception:
                    uploaded_file.seek(0)
                    df_upload = pd.read_csv(uploaded_file)
            else:
                df_upload = pd.read_excel(uploaded_file)
            
            # Normalisasi nama kolom agar tidak sensitif huruf besar/kecil & spasi
            column_map = {}
            for col in df_upload.columns:
                c_clean = str(col).strip().upper()
                if c_clean in ["NO", "NO.", "NOMOR"]:
                    column_map[col] = "No"
                elif "NIS" in c_clean or "INDUK" in c_clean or "NISN" in c_clean:
                    column_map[col] = "Nomor Induk / NISN"
                elif "NAMA" in c_clean:
                    column_map[col] = "Nama Siswa"
                elif c_clean in ["JK", "JENIS KELAMIN", "GENDER", "L/P", "SEX"]:
                    column_map[col] = "Jenis Kelamin"
                elif "HBE" in c_clean or "EFEKTIF" in c_clean:
                    column_map[col] = "HBE"
                elif "SAKIT" in c_clean or c_clean == "S":
                    column_map[col] = "Sakit"
                elif "IZIN" in c_clean or "IJIN" in c_clean or c_clean == "I":
                    column_map[col] = "Izin"
                elif "ALPA" in c_clean or "ALPHA" in c_clean or "ABSEN" in c_clean or c_clean == "A":
                    column_map[col] = "Alpa"

            df_upload = df_upload.rename(columns=column_map)

            # Buat kolom otomatis jika tidak ada di CSV
            if "No" not in df_upload.columns:
                df_upload["No"] = list(range(1, len(df_upload) + 1))
            if "Nomor Induk / NISN" not in df_upload.columns:
                df_upload["Nomor Induk / NISN"] = "-"
            if "Nama Siswa" not in df_upload.columns:
                # Ambil kolom pertama sebagai nama jika tidak terdeteksi
                df_upload["Nama Siswa"] = df_upload.iloc[:, 0]
            if "Jenis Kelamin" not in df_upload.columns:
                df_upload["Jenis Kelamin"] = "L"
            if "HBE" not in df_upload.columns:
                df_upload["HBE"] = 24
            if "Sakit" not in df_upload.columns:
                df_upload["Sakit"] = 0
            if "Izin" not in df_upload.columns:
                df_upload["Izin"] = 0
            if "Alpa" not in df_upload.columns:
                df_upload["Alpa"] = 0

            # Format kolom Jenis Kelamin
            df_upload["Jenis Kelamin"] = df_upload["Jenis Kelamin"].astype(str).str.strip().str.upper()
            df_upload["Jenis Kelamin"] = df_upload["Jenis Kelamin"].apply(lambda x: "P" if x in ["P", "PEREMPUAN", "FEMALE"] else "L")

            required_cols = ["No", "Nomor Induk / NISN", "Nama Siswa", "Jenis Kelamin", "HBE", "Sakit", "Izin", "Alpa"]
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
                row.get("Izin", 0),
                row.get("Alpa", 0),
                row.get("Jumlah Ketidakhadiran", 0),
                row.get("Jumlah Hadir", 0),
                f"=(J{curr_row}/E{curr_row})"
            ])
            ws.cell(row=curr_row, column=11).number_format = '0.0%'

        last_row = ws.max_row
        total_row = last_row + 1
        
        ws.append([
            "JUMLAH TOTAL", "", "", "",
            f"=SUM(E{start_data_row}:E{last_row})",
            f"=SUM(F{start_data_row}:F{last_row})",
            f"=SUM(G{start_data_row}:G{last_row})",
            f"=SUM(H{start_data_row}:H{last_row})",
            f"=SUM(I{start_data_row}:I{last_row})",
            f"=SUM(J{start_data_row}:J{last_row})",
            f"=J{total_row}/E{total_row}"
        ])
        ws.cell(row=total_row, column=11).number_format = '0.0%'

        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=4)

        apply_excel_styling(
            ws, headers,
            nama_sekolah.upper(),
            f"LAPORAN REKAPITULASI ABSENSI SISWA - {kelas.upper()}",
            f"PERIODE: {bulan_tahun.upper()}",
            wali_kelas,
            tgl_cetak_str,
            logo_file=uploaded_logo
        )

        wb.save(output)
        return output.getvalue()

    st.download_button(
        label="📥 Download Excel Rekap Bulanan",
        data=generate_excel_rekap_bulanan(),
        file_name=f"Rekap_Absensi_Bulanan_{kelas}_{bulan_tahun}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# TAB 2: KEHADIRAN SISWA PER HARI
# ==========================================
with tab2:
    st.subheader("📅 Rekap Kehadiran Harian Siswa Per Tanggal")
    tgl_pilih = st.date_input("Pilih Tanggal Absensi", key="tgl_harian")

    if "data_harian" not in st.session_state:
        st.session_state.data_harian = pd.DataFrame({
            "No": [1, 2, 3],
            "Nomor Induk": ["1001", "1002", "1003"],
            "Nama Siswa": ["Ahmad Fauzi", "Budi Santoso", "Citra Dewi"],
            "Status Kehadiran": ["Hadir", "Hadir", "Izin"],
            "Keterangan": ["-", "-", "Acara Keluarga"]
        })

    edited_harian = st.data_editor(
        st.session_state.data_harian,
        column_config={
            "Status Kehadiran": st.column_config.SelectboxColumn(
                "Status Kehadiran",
                options=["Hadir", "Sakit", "Izin", "Alpa"],
                required=True
            )
        },
        num_rows="dynamic",
        use_container_width=True,
        key="editor_harian"
    )

    total_siswa = len(edited_harian)
    hadir_count = (edited_harian["Status Kehadiran"] == "Hadir").sum()
    sakit_count = (edited_harian["Status Kehadiran"] == "Sakit").sum()
    izin_count = (edited_harian["Status Kehadiran"] == "Izin").sum()
    alpa_count = (edited_harian["Status Kehadiran"] == "Alpa").sum()
    persen_hadir = (hadir_count / total_siswa * 100) if total_siswa > 0 else 0

    col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns(5)
    col_h1.metric("Hadir", f"{hadir_count} Siswa")
    col_h2.metric("Sakit", f"{sakit_count} Siswa")
    col_h3.metric("Izin", f"{izin_count} Siswa")
    col_h4.metric("Alpa", f"{alpa_count} Siswa")
    col_h5.metric("% Kehadiran Hari Ini", f"{persen_hadir:.1f}%")

    def generate_excel_harian():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Kehadiran Harian"

        headers = ["No", "Nomor Induk", "Nama Siswa", "Status Kehadiran", "Keterangan"]
        
        ws.append([nama_sekolah.upper()])
        ws.append([f"LAPORAN KEHADIRAN HARIAN SISWA - {kelas.upper()}"])
        ws.append([f"TANGGAL: {tgl_pilih.strftime('%d-%m-%Y')}"])
        ws.append([])
        ws.append(headers)

        for _, row in edited_harian.iterrows():
            ws.append([
                row.get("No", ""),
                row.get("Nomor Induk", ""),
                row.get("Nama Siswa", ""),
                row.get("Status Kehadiran", ""),
                row.get("Keterangan", "-")
            ])

        ws.append([])
        ws.append(["RINGKASAN KEHADIRAN HARIAN"])
        ws.append(["Total Siswa", total_siswa])
        ws.append(["Total Hadir", hadir_count])
        ws.append(["Total Sakit", sakit_count])
        ws.append(["Total Izin", izin_count])
        ws.append(["Total Alpa", alpa_count])
        ws.append(["Persentase Kehadiran", f"{persen_hadir:.1f}%"])

        apply_excel_styling(
            ws, headers,
            nama_sekolah.upper(),
            f"LAPORAN KEHADIRAN HARIAN SISWA - {kelas.upper()}",
            f"TANGGAL: {tgl_pilih.strftime('%d-%m-%Y')}",
            wali_kelas,
            tgl_cetak_str,
            logo_file=uploaded_logo
        )

        wb.save(output)
        return output.getvalue()

    st.download_button(
        label="📥 Download Excel Kehadiran Harian",
        data=generate_excel_harian(),
        file_name=f"Kehadiran_Harian_{kelas}_{tgl_pilih.strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# TAB 3: MUTASI SISWA (MASUK / KELUAR)
# ==========================================
with tab3:
    st.subheader("🔄 Catatan Mutasi Siswa (Masuk / Keluar)")

    if "data_mutasi" not in st.session_state:
        st.session_state.data_mutasi = pd.DataFrame({
            "No": [1],
            "Tanggal Mutasi": ["2026-10-01"],
            "Nomor Induk / NISN": ["1004"],
            "Nama Siswa": ["Eko Prasetyo"],
            "Jenis Kelamin": ["L"],
            "Jenis Mutasi": ["Masuk"],
            "Asal / Tujuan Sekolah": ["SMP Negeri 2"],
            "Alasan Mutasi": ["Pindah Tugas Orang Tua"]
        })

    edited_mutasi = st.data_editor(
        st.session_state.data_mutasi,
        column_config={
            "Jenis Kelamin": st.column_config.SelectboxColumn("Jenis Kelamin", options=["L", "P"], required=True),
            "Jenis Mutasi": st.column_config.SelectboxColumn("Jenis Mutasi", options=["Masuk", "Keluar"], required=True)
        },
        num_rows="dynamic",
        use_container_width=True,
        key="editor_mutasi"
    )

    m_masuk = (edited_mutasi["Jenis Mutasi"] == "Masuk").sum()
    m_keluar = (edited_mutasi["Jenis Mutasi"] == "Keluar").sum()
    
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Total Siswa Masuk", f"{m_masuk} Orang")
    col_m2.metric("Total Siswa Keluar", f"{m_keluar} Orang")

    def generate_excel_mutasi():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Mutasi Siswa"

        headers = ["No", "Tanggal Mutasi", "Nomor Induk", "Nama Siswa", "JK", "Jenis Mutasi", "Asal / Tujuan Sekolah", "Alasan Mutasi"]
        
        ws.append([nama_sekolah.upper()])
        ws.append([f"LAPORAN MUTASI SISWA (MASUK / KELUAR) - {kelas.upper()}"])
        ws.append([f"PERIODE: {bulan_tahun.upper()}"])
        ws.append([])
        ws.append(headers)

        for _, row in edited_mutasi.iterrows():
            ws.append([
                row.get("No", ""),
                str(row.get("Tanggal Mutasi", "")),
                row.get("Nomor Induk / NISN", ""),
                row.get("Nama Siswa", ""),
                row.get("Jenis Kelamin", ""),
                row.get("Jenis Mutasi", ""),
                row.get("Asal / Tujuan Sekolah", ""),
                row.get("Alasan
