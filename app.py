import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from datetime import date
from PIL import Image

# ==========================================
# 0. KONFIGURASI PENYIMPANAN DATA PERMANEN
# ==========================================
DATA_FILE = "database_absensi_siswa.csv"

# Data Default jika file database belum ada
DATA_DEFAULT = pd.DataFrame([
    {"No": 1, "Nama Murid": "ABANG MUSHAWIR EDO", "L/P": "L", "Nomor Induk": "0012345678", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 2, "Nama Murid": "AHMAD YANI", "L/P": "L", "Nomor Induk": "0012345679", "HBE": 25, "S": 0, "I": 0, "A": 2},
    {"No": 3, "Nama Murid": "AL JAMI", "L/P": "L", "Nomor Induk": "0012345680", "HBE": 25, "S": 1, "I": 0, "A": 3},
    {"No": 4, "Nama Murid": "AYU NINGSIH", "L/P": "P", "Nomor Induk": "0012345681", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 5, "Nama Murid": "EPRI SASKIA", "L/P": "P", "Nomor Induk": "0012345682", "HBE": 25, "S": 0, "I": 0, "A": 2},
])

def load_saved_data():
    """Membaca data dari file lokal jika ada."""
    if os.path.exists(DATA_FILE):
        try:
            df = pd.read_csv(DATA_FILE, dtype={"Nomor Induk": str})
            df["No"] = range(1, len(df) + 1)
            return df
        except Exception:
            return DATA_DEFAULT.copy()
    else:
        return DATA_DEFAULT.copy()

def save_data_permanently(df):
    """Menyimpan data ke file lokal secara permanen."""
    df_to_save = df.copy()
    df_to_save["No"] = range(1, len(df_to_save) + 1)
    df_to_save.to_csv(DATA_FILE, index=False)


# ==========================================
# 1. KONFIGURASI HALAMAN & STATE
# ==========================================
st.set_page_config(page_title="Rekapitulasi Absensi Siswa", layout="wide")

# Load data tersimpan saat pertama kali aplikasi dijalankan / di-refresh
if "data_absensi" not in st.session_state:
    st.session_state.data_absensi = load_saved_data()

# ==========================================
# 2. SIDEBAR - INFORMASI & FITUR UPLOAD
# ==========================================
st.sidebar.header("🏫 Data Sekolah & Kelas")

uploaded_logo = st.sidebar.file_uploader("Upload Logo Sekolah (PNG / JPG)", type=["png", "jpg", "jpeg"])
logo_img = None
if uploaded_logo is not None:
    try:
        logo_img = Image.open(uploaded_logo)
        st.sidebar.image(logo_img, width=100, caption="Logo Sekolah")
    except Exception:
        pass

nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SMP NEGERI 1 NANGA MAHAP")
tahun_pelajaran = st.sidebar.text_input("Tahun Pelajaran", "2025/2026")
kelas = st.sidebar.text_input("Kelas", "IX C")
bulan_tahun = st.sidebar.text_input("Bulan / Periode", "April 2026")
tempat_cetak = st.sidebar.text_input("Kota / Tempat Laporan", "Nanga Mahap")
tgl_cetak = st.sidebar.date_input("Tanggal Cetak Laporan", date(2026, 4, 30))
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "Triyanto trisno ,S.Pd")
nip_wali = st.sidebar.text_input("NIP Wali Kelas", "199305202024211001")

bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
tgl_cetak_str = f"{tempat_cetak}, {tgl_cetak.day} {bulan_indo[tgl_cetak.month - 1]} {tgl_cetak.year}"

# ------------------------------------------
# FITUR UPLOAD DATA SISWA (DETEKSI SEPARATOR CSV AUTO)
# ------------------------------------------
st.sidebar.divider()
st.sidebar.header("📁 Upload File Data Siswa")
st.sidebar.caption("Format file yang didukung: Excel (.xlsx) atau CSV (.csv) berisi kolom `NOMOR`, `NAMA SISWA`, `L/P`, `NISN`")

uploaded_data_file = st.sidebar.file_uploader("Upload File Siswa (.xlsx / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_data_file is not None:
    if st.sidebar.button("📥 Terapkan Data dari File", type="primary"):
        try:
            # Baca file dengan auto deteksi separator (; atau ,)
            if uploaded_data_file.name.endswith(".csv"):
                try:
                    df_raw = pd.read_csv(uploaded_data_file, sep=None, engine='python')
                except Exception:
                    uploaded_data_file.seek(0)
                    df_raw = pd.read_csv(uploaded_data_file, sep=';')
            else:
                df_raw = pd.read_excel(uploaded_data_file)
            
            df_raw = df_raw.dropna(how="all")
            
            col_map = {str(col).strip().upper(): col for col in df_raw.columns}

            def find_col(keywords):
                for kw in keywords:
                    for clean_key, orig_key in col_map.items():
                        if kw in clean_key:
                            return df_raw[orig_key]
                return None

            series_nama = find_col(["NAMA SISWA", "NAMA MURID", "NAMA", "SISWA"])
            series_nisn = find_col(["NISN", "NOMOR INDUK", "NIS", "INDUK", "NO INDUK"])
            series_lp = find_col(["L/P", "JK", "JENIS KELAMIN", "KELAMIN", "SEX"])
            series_hbe = find_col(["HBE", "HARI BELAJAR"])
            series_s = find_col(["S", "SAKIT"])
            series_i = find_col(["I", "IZIN"])
            series_a = find_col(["A", "ALPHA", "ALPA"])

            # Jika header tidak cocok, gunakan posisi kolom
            if series_nama is None and len(df_raw.columns) >= 2:
                series_nama = df_raw.iloc[:, 1]
                if series_nisn is None and len(df_raw.columns) >= 4:
                    series_nisn = df_raw.iloc[:, 3]
                if series_lp is None and len(df_raw.columns) >= 3:
                    series_lp = df_raw.iloc[:, 2]

            if series_nama is None:
                st.sidebar.error("❌ Gagal membaca nama siswa. Pastikan file CSV/Excel memiliki kolom NAMA SISWA.")
            else:
                df_new = pd.DataFrame()
                df_new["Nama Murid"] = series_nama.astype(str).str.strip().str.upper()
                df_new = df_new[df_new["Nama Murid"].notna() & (df_new["Nama Murid"] != "NAN") & (df_new["Nama Murid"] != "")]
                
                if series_nisn is not None:
                    df_new["Nomor Induk"] = series_nisn.astype(str).str.replace(".0", "", regex=False).str.strip().replace("nan", "")
                else:
                    df_new["Nomor Induk"] = ""

                if series_lp is not None:
                    df_new["L/P"] = series_lp.astype(str).str.strip().str.upper().apply(
                        lambda x: "P" if x in ["P", "PEREMPUAN", "FEMALE"] else "L"
                    )
                else:
                    df_new["L/P"] = "L"

                df_new["HBE"] = pd.to_numeric(series_hbe, errors="coerce").fillna(25).astype(int) if series_hbe is not None else 25
                df_new["S"] = pd.to_numeric(series_s, errors="coerce").fillna(0).astype(int) if series_s is not None else 0
                df_new["I"] = pd.to_numeric(series_i, errors="coerce").fillna(0).astype(int) if series_i is not None else 0
                df_new["A"] = pd.to_numeric(series_a, errors="coerce").fillna(0).astype(int) if series_a is not None else 0
                
                df_new["No"] = range(1, len(df_new) + 1)
                
                cols_order = ["No", "Nama Murid", "L/P", "Nomor Induk", "HBE", "S", "I", "A"]
                final_df = df_new[cols_order].reset_index(drop=True).copy()
                
                # Simpan ke session state DAN simpan permanen ke file
                st.session_state.data_absensi = final_df
                save_data_permanently(final_df)
                
                st.sidebar.success(f"✅ Berhasil memproses & menyimpan {len(final_df)} data siswa!")
                st.rerun()

        except Exception as e:
            st.sidebar.error(f"Gagal membaca file: {e}")

# ------------------------------------------
# FITUR TAMBAH SISWA MANUAL
# ------------------------------------------
st.sidebar.divider()
st.sidebar.header("➕ Tambah Siswa Manual")

with st.sidebar.form("form_tambah_siswa", clear_on_submit=True):
    input_nama = st.text_input("Nama Lengkap Siswa")
    input_lp = st.selectbox("Jenis Kelamin (L/P)", ["L", "P"])
    input_nis = st.text_input("Nomor Induk / NISN", "")
    input_hbe = st.number_input("Hari Belajar Efektif (HBE)", min_value=1, value=25)
    
    btn_tambah = st.form_submit_button("➕ Tambahkan Siswa")

    if btn_tambah:
        if input_nama.strip() == "":
            st.sidebar.error("Nama siswa tidak boleh kosong!")
        else:
            no_baru = len(st.session_state.data_absensi) + 1
            siswa_baru = {
                "No": no_baru,
                "Nama Murid": input_nama.upper().strip(),
                "L/P": input_lp,
                "Nomor Induk": input_nis.strip(),
                "HBE": input_hbe,
                "S": 0,
                "I": 0,
                "A": 0
            }
            new_df = pd.concat(
                [st.session_state.data_absensi, pd.DataFrame([siswa_baru])], 
                ignore_index=True
            )
            st.session_state.data_absensi = new_df
            save_data_permanently(new_df)
            st.sidebar.success(f"Berhasil menambahkan & menyimpan {input_nama.upper()}!")
            st.rerun()

# Tombol Reset ke Data Bawaan
st.sidebar.divider()
if st.sidebar.button("🔄 Reset ke Data Default"):
    st.session_state.data_absensi = DATA_DEFAULT.copy()
    save_data_permanently(DATA_DEFAULT)
    st.sidebar.info("Data telah dikembalikan ke standar awal.")
    st.rerun()

# ==========================================
# 3. HEADER LAPORAN UTAMA
# ==========================================
st.markdown("<h2 style='text-align: center; margin-bottom: 0px;'>REKAPITULASI ABSENSI SISWA</h2>", unsafe_allow_html=True)
st.markdown(f"<h3 style='text-align: center; margin-top: 0px; margin-bottom: 0px;'>{nama_sekolah.upper()}</h3>", unsafe_allow_html=True)
st.markdown(f"<h4 style='text-align: center; margin-top: 0px;'>TAHUN PELAJARAN {tahun_pelajaran}</h4>", unsafe_allow_html=True)

st.markdown(f"**KELAS : {kelas.upper()}**")

# ==========================================
# 4. INPUT / EDIT TABEL INTERAKTIF
# ==========================================
st.markdown("### 📝 Input / Edit Data Absensi Siswa")
st.caption("Ubah isi tabel langsung di bawah ini. Tekan tombol **💾 Simpan Perubahan Data** untuk menyimpan secara permanen agar tidak hilang saat di-refresh.")

df_to_edit = st.session_state.data_absensi.copy().reset_index(drop=True)

edited_df = st.data_editor(
    df_to_edit,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="tabel_absensi_v6"
)

# Tombol Simpan Perubahan Edit Tabel
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    if st.button("💾 Simpan Perubahan Data", type="primary"):
        edited_df["No"] = range(1, len(edited_df) + 1)
        st.session_state.data_absensi = edited_df.reset_index(drop=True).copy()
        save_data_permanently(st.session_state.data_absensi)
        st.success("✅ Perubahan data berhasil disimpan secara permanen!")

# ==========================================
# 5. KALKULASI LAPORAN
# ==========================================
df_calc = st.session_state.data_absensi.copy()
for col in ["HBE", "S", "I", "A"]:
    df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)

df_calc["JUMLAH"] = df_calc["S"] + df_calc["I"] + df_calc["A"]

# Kalkulasi Presentase
df_calc["PRESENTASE S"] = (df_calc["S"] / df_calc["HBE"]).fillna(0)
df_calc["PRESENTASE I"] = (df_calc["I"] / df_calc["HBE"]).fillna(0)
df_calc["PRESENTASE A"] = (df_calc["A"] / df_calc["HBE"]).fillna(0)
df_calc["PRESENTASE KEHADIRAN %"] = ((df_calc["HBE"] - df_calc["JUMLAH"]) / df_calc["HBE"]).fillna(0)

# Tampilan Tabel Web
df_view = pd.DataFrame()
df_view["NO"] = df_calc["No"].astype(str)
df_view["NAMA MURID"] = df_calc["Nama Murid"]
df_view["L/P"] = df_calc["L/P"]
df_view["NOMOR INDUK"] = df_calc["Nomor Induk"]
df_view["HBE"] = df_calc["HBE"].astype(str)
df_view["ABSENSI (S)"] = df_calc["S"].astype(str)
df_view["ABSENSI (I)"] = df_calc["I"].astype(str)
df_view["ABSENSI (A)"] = df_calc["A"].astype(str)
df_view["JUMLAH"] = df_calc["JUMLAH"].astype(str)
df_view["PRESENTASE (S)"] = (df_calc["PRESENTASE S"] * 100).round(0).astype(int).astype(str) + "%"
df_view["PRESENTASE (I)"] = (df_calc["PRESENTASE I"] * 100).round(0).astype(int).astype(str) + "%"
df_view["PRESENTASE (A)"] = (df_calc["PRESENTASE A"] * 100).round(0).astype(int).astype(str) + "%"
df_view["PRESENTASE KEHADIRAN %"] = (df_calc["PRESENTASE KEHADIRAN %"] * 100).round(0).astype(int).astype(str) + "%"

# Baris Jumlah Total
total_l = (df_calc["L/P"] == "L").sum()
total_p = (df_calc["L/P"] == "P").sum()
total_siswa = len(df_calc)

total_s = df_calc["S"].sum()
total_i = df_calc["I"].sum()
total_a = df_calc["A"].sum()
total_jumlah_absen = df_calc["JUMLAH"].sum()

avg_s_pct = f"{round(df_calc['PRESENTASE S'].mean() * 100) if len(df_calc) > 0 else 0}%"
avg_i_pct = f"{round(df_calc['PRESENTASE I'].mean() * 100) if len(df_calc) > 0 else 0}%"
avg_a_pct = f"{round(df_calc['PRESENTASE A'].mean() * 100) if len(df_calc) > 0 else 0}%"
avg_hadir_pct = f"{round(df_calc['PRESENTASE KEHADIRAN %'].mean() * 100) if len(df_calc) > 0 else 0}%"

total_row = {
    "NO": "",
    "NAMA MURID": "JUMLAH",
    "L/P": str(total_siswa),
    "NOMOR INDUK": "",
    "HBE": "",
    "ABSENSI (S)": str(total_s),
    "ABSENSI (I)": str(total_i),
    "ABSENSI (A)": str(total_a),
    "JUMLAH": str(total_jumlah_absen),
    "PRESENTASE (S)": avg_s_pct,
    "PRESENTASE (I)": avg_i_pct,
    "PRESENTASE (A)": avg_a_pct,
    "PRESENTASE KEHADIRAN %": avg_hadir_pct,
}

df_view_with_total = pd.concat([df_view, pd.DataFrame([total_row])], ignore_index=True)

st.markdown("### 📊 Hasil Rekapitulasi Absensi Siswa")
st.dataframe(df_view_with_total, use_container_width=True, hide_index=True)

st.markdown(f"""
**Ringkasan Siswa:**
- **Laki - Laki** : {total_l}
- **Perempuan** : {total_p}
- **Jumlah akhir bulan** : {total_siswa}
""")

# ==========================================
# 6. EXCEL GENERATOR (SIAP CETAK)
# ==========================================
def generate_excel_laporan():
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rekap Absensi"

    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True)
    font_subtitle = Font(name="Calibri", size=11, bold=True)
    font_header = Font(name="Calibri", size=10, bold=True)
    font_data = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    if uploaded_logo is not None and logo_img is not None:
        try:
            uploaded_logo.seek(0)
            img_for_excel = OpenpyxlImage(uploaded_logo)
            img_for_excel.height = 65
            img_for_excel.width = int(65 * (logo_img.width / logo_img.height)) if logo_img.height else 65
            ws.add_image(img_for_excel, "A1")
        except Exception:
            pass

    ws.merge_cells("A1:M1")
    ws.cell(row=1, column=1, value="REKAPITULASI ABSENSI SISWA").font = font_title
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:M2")
    ws.cell(row=2, column=1, value=nama_sekolah.upper()).font = font_title
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A3:M3")
    ws.cell(row=3, column=1, value=f"TAHUN PELAJARAN {tahun_pelajaran}").font = font_subtitle
    ws.cell(row=3, column=1).alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=5, column=1, value=f"KELAS   : {kelas.upper()}").font = font_bold

    headers_r7 = [
        ("NO", "A7", "A8"),
        ("NAMA MURID", "B7", "B8"),
        ("L/P", "C7", "C8"),
        ("NOMOR INDUK", "D7", "D8"),
        ("HBE", "E7", "E8"),
        ("ABSENSI", "F7", "H7"),
        ("JUMLAH", "I7", "I8"),
        ("PRESENTASE", "J7", "L7"),
        ("PRESENTASE KEHADIRAN %", "M7", "M8")
    ]

    for title, start_col, end_col in headers_r7:
        if start_col != end_col:
            ws.merge_cells(f"{start_col}:{end_col}")
        cell = ws[start_col.split(":")[0]]
        cell.value = title
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.cell(row=8, column=6, value="S").font = font_header
    ws.cell(row=8, column=7, value="I").font = font_header
    ws.cell(row=8, column=8, value="A").font = font_header

    ws.cell(row=8, column=10, value="S").font = font_header
    ws.cell(row=8, column=11, value="I").font = font_header
    ws.cell(row=8, column=12, value="A").font = font_header

    for c in range(1, 14):
        ws.cell(row=7, column=c).border = thin_border
        ws.cell(row=8, column=c).border = thin_border
        ws.cell(row=8, column=c).alignment = Alignment(horizontal="center", vertical="center")

    start_row = 9
    num_students = len(df_calc)
    for idx, row in df_calc.iterrows():
        r = start_row + idx
        ws.cell(row=r, column=1, value=idx + 1)
        ws.cell(row=r, column=2, value=row.get("Nama Murid", ""))
        
        lp_cell = ws.cell(row=r, column=3, value=row.get("L/P", "L"))
        lp_cell.fill = green_fill
        
        ws.cell(row=r, column=4, value=row.get("Nomor Induk", ""))
        ws.cell(row=r, column=5, value=row.get("HBE", 25))
        ws.cell(row=r, column=6, value=row.get("S", 0))
        ws.cell(row=r, column=7, value=row.get("I", 0))
        ws.cell(row=r, column=8, value=row.get("A", 0))
        
        ws.cell(row=r, column=9, value=f"=SUM(F{r}:H{r})")
        ws.cell(row=r, column=10, value=f"=F{r}/E{r}").number_format = '0%'
        ws.cell(row=r, column=11, value=f"=G{r}/E{r}").number_format = '0%'
        ws.cell(row=r, column=12, value=f"=H{r}/E{r}").number_format = '0%'
        ws.cell(row=r, column=13, value=f"=(E{r}-I{r})/E{r}").number_format = '0%'

        for c in range(1, 14):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.font = font_bold if c in [1, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13] else font_data
            if c in [2, 4]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    total_row_idx = start_row + num_students
    ws.merge_cells(f"A{total_row_idx}:B{total_row_idx}")
    cell_tot_lbl = ws.cell(row=total_row_idx, column=1, value="JUMLAH")
    cell_tot_lbl.font = font_title
    cell_tot_lbl.alignment = Alignment(horizontal="center", vertical="center")

    if num_students > 0:
        ws.cell(row=total_row_idx, column=3, value=f"=COUNTA(C9:C{total_row_idx-1})")
        ws.cell(row=total_row_idx, column=6, value=f"=SUM(F9:F{total_row_idx-1})")
        ws.cell(row=total_row_idx, column=7, value=f"=SUM(G9:G{total_row_idx-1})")
        ws.cell(row=total_row_idx, column=8, value=f"=SUM(H9:H{total_row_idx-1})")
        ws.cell(row=total_row_idx, column=9, value=f"=SUM(I9:I{total_row_idx-1})")
        
        ws.cell(row=total_row_idx, column=10, value=f"=AVERAGE(J9:J{total_row_idx-1})").number_format = '0%'
        ws.cell(row=total_row_idx, column=11, value=f"=AVERAGE(K9:K{total_row_idx-1})").number_format = '0%'
        ws.cell(row=total_row_idx, column=12, value=f"=AVERAGE(L9:L{total_row_idx-1})").number_format = '0%'
        ws.cell(row=total_row_idx, column=13, value=f"=AVERAGE(M9:M{total_row_idx-1})").number_format = '0%'

    for c in range(1, 14):
        cell = ws.cell(row=total_row_idx, column=c)
        cell.fill = yellow_fill
        cell.border = thin_border
        cell.font = font_bold
        if c not in [1, 2]:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    r_sum1 = total_row_idx + 2
    r_sum2 = r_sum1 + 1
    r_sum3 = r_sum2 + 1

    ws.cell(row=r_sum1, column=1, value="Laki – Laki").font = font_bold
    ws.cell(row=r_sum1, column=3, value=":").alignment = Alignment(horizontal="center")
    ws.cell(row=r_sum1, column=4, value=f'=COUNTIF(C9:C{total_row_idx-1}, "L")' if num_students > 0 else 0).font = font_bold

    ws.cell(row=r_sum2, column=1, value="Perempuan").font = font_bold
    ws.cell(row=r_sum2, column=3, value=":").alignment = Alignment(horizontal="center")
    ws.cell(row=r_sum2, column=4, value=f'=COUNTIF(C9:C{total_row_idx-1}, "P")' if num_students > 0 else 0).font = font_bold

    ws.cell(row=r_sum3, column=1, value="Jumlah akhir bulan").font = font_bold
    ws.cell(row=r_sum3, column=3, value=":").alignment = Alignment(horizontal="center")
    ws.cell(row=r_sum3, column=4, value=f'=COUNTA(C9:C{total_row_idx-1})' if num_students > 0 else 0).font = font_bold

    r_ttd_tgl = total_row_idx + 4
    r_ttd_jab = r_ttd_tgl + 1
    r_ttd_nama = r_ttd_jab + 4
    r_ttd_nip = r_ttd_nama + 1

    ws.cell(row=r_ttd_tgl, column=10, value=tgl_cetak_str).font = font_data
    ws.cell(row=r_ttd_jab, column=10, value=f"Wali Kelas {kelas}").font = font_data
    ws.cell(row=r_ttd_nama, column=10, value=wali_kelas).font = Font(name="Calibri", size=10, bold=True, underline="single")
    ws.cell(row=r_ttd_nip, column=10, value=f"NIP. {nip_wali}" if nip_wali else "").font = font_bold

    col_widths = {
        'A': 6, 'B': 30, 'C': 6, 'D': 18, 'E': 6,
        'F': 5, 'G': 5, 'H': 5, 'I': 8,
        'J': 8, 'K': 8, 'L': 8, 'M': 16
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    wb.save(output)
    return output.getvalue()

st.download_button(
    label="📥 Download Excel Rekap Absensi (Siap Cetak)",
    data=generate_excel_laporan(),
    file_name=f"Rekap_Absensi_Siswa_{kelas}_{tahun_pelajaran.replace('/', '-')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
