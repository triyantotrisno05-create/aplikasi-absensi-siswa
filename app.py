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
FILE_ABSENSI = "database_absensi_siswa.csv"
FILE_HARIAN = "database_kehadiran_harian.csv"
FILE_MUTASI = "database_mutasi_siswa.csv"

# Data Default Absensi
DEFAULT_ABSENSI = pd.DataFrame([
    {"No": 1, "Nama Murid": "ABANG MUSHAWIR EDO", "L/P": "L", "Nomor Induk": "0012345678", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 2, "Nama Murid": "AHMAD YANI", "L/P": "L", "Nomor Induk": "0012345679", "HBE": 25, "S": 0, "I": 0, "A": 2},
    {"No": 3, "Nama Murid": "AL JAMI", "L/P": "L", "Nomor Induk": "0012345680", "HBE": 25, "S": 1, "I": 0, "A": 3},
    {"No": 4, "Nama Murid": "AYU NINGSIH", "L/P": "P", "Nomor Induk": "0012345681", "HBE": 25, "S": 0, "I": 0, "A": 0},
    {"No": 5, "Nama Murid": "EPRI SASKIA", "L/P": "P", "Nomor Induk": "0012345682", "HBE": 25, "S": 0, "I": 0, "A": 2},
])

# Data Default Kehadiran Harian
DEFAULT_HARIAN = pd.DataFrame([
    {"Tanggal": str(date.today()), "Hadir": 5, "Sakit": 0, "Izin": 0, "Alpha": 0, "Total Siswa": 5, "Persentase Kehadiran (%)": 100.0}
])

# Data Default Mutasi
DEFAULT_MUTASI = pd.DataFrame([
    {"Tanggal Mutasi": str(date.today()), "Nama Siswa": "BUDI SANTOSO", "L/P": "L", "NISN": "0019998877", "Jenis Mutasi": "Masuk", "Keterangan / Sekolah Asal/Tujuan": "SMPN 2 Nanga Mahap"}
])

def load_data(filepath, default_df):
    if os.path.exists(filepath):
        try:
            return pd.read_csv(filepath)
        except Exception:
            return default_df.copy()
    return default_df.copy()

def save_data(df, filepath):
    df.to_csv(filepath, index=False)

# ==========================================
# 1. KONFIGURASI HALAMAN & STATE
# ==========================================
st.set_page_config(page_title="Sistem Informasi Absensi & Mutasi Siswa", layout="wide")

if "data_absensi" not in st.session_state:
    st.session_state.data_absensi = load_data(FILE_ABSENSI, DEFAULT_ABSENSI)

if "data_harian" not in st.session_state:
    st.session_state.data_harian = load_data(FILE_HARIAN, DEFAULT_HARIAN)

if "data_mutasi" not in st.session_state:
    st.session_state.data_mutasi = load_data(FILE_MUTASI, DEFAULT_MUTASI)

# ==========================================
# 2. SIDEBAR & INFORMASI SEKOLAH
# ==========================================
st.sidebar.header("🏫 Data Sekolah & Kelas")

uploaded_logo = st.sidebar.file_uploader("Upload Logo Sekolah (PNG/JPG)", type=["png", "jpg", "jpeg"])
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
tgl_cetak = st.sidebar.date_input("Tanggal Cetak Laporan", date.today())
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "Triyanto trisno, S.Pd")
nip_wali = st.sidebar.text_input("NIP Wali Kelas", "199305202024211001")

# UPLOAD FILE DATA SISWA
st.sidebar.divider()
st.sidebar.header("📁 Upload File Data Siswa")
uploaded_data_file = st.sidebar.file_uploader("Upload File Siswa (.xlsx / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_data_file is not None:
    if st.sidebar.button("📥 Terapkan Data dari File", type="primary"):
        try:
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
            series_nisn = find_col(["NISN", "NOMOR INDUK", "NIS", "INDUK"])
            series_lp = find_col(["L/P", "JK", "JENIS KELAMIN", "KELAMIN"])

            if series_nama is None and len(df_raw.columns) >= 2:
                series_nama = df_raw.iloc[:, 1]
                if series_nisn is None and len(df_raw.columns) >= 4:
                    series_nisn = df_raw.iloc[:, 3]
                if series_lp is None and len(df_raw.columns) >= 3:
                    series_lp = df_raw.iloc[:, 2]

            if series_nama is not None:
                df_new = pd.DataFrame()
                df_new["Nama Murid"] = series_nama.astype(str).str.strip().str.upper()
                df_new = df_new[df_new["Nama Murid"].notna() & (df_new["Nama Murid"] != "NAN") & (df_new["Nama Murid"] != "")]
                
                df_new["Nomor Induk"] = series_nisn.astype(str).str.replace(".0", "", regex=False).str.strip().replace("nan", "") if series_nisn is not None else ""
                df_new["L/P"] = series_lp.astype(str).str.strip().str.upper().apply(lambda x: "P" if x in ["P", "PEREMPUAN"] else "L") if series_lp is not None else "L"
                df_new["HBE"] = 25
                df_new["S"] = 0
                df_new["I"] = 0
                df_new["A"] = 0
                df_new["No"] = range(1, len(df_new) + 1)
                
                cols_order = ["No", "Nama Murid", "L/P", "Nomor Induk", "HBE", "S", "I", "A"]
                final_df = df_new[cols_order].reset_index(drop=True)
                st.session_state.data_absensi = final_df
                save_data(final_df, FILE_ABSENSI)
                st.sidebar.success(f"✅ Berhasil mengimpor {len(final_df)} siswa!")
                st.rerun()
            else:
                st.sidebar.error("❌ Kolom nama siswa tidak ditemukan.")
        except Exception as e:
            st.sidebar.error(f"Gagal membaca file: {e}")

# ==========================================
# 3. NAVIGASI TAB UTAMA
# ==========================================
st.markdown(f"<h2 style='text-align: center;'>{nama_sekolah.upper()}</h2>", unsafe_allow_html=True)
st.markdown(f"<h4 style='text-align: center;'>REKAPITULASI ABSENSI & MUTASI SISWA - {kelas.upper()} ({tahun_pelajaran})</h4>", unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs([
    "📊 Rekap Absensi Bulanan", 
    "📅 Persentase Kehadiran Per Hari", 
    "🔄 Mutasi Siswa"
])

# ------------------------------------------
# TAB 1: REKAP ABSENSI BULANAN
# ------------------------------------------
with tab1:
    st.subheader("📝 Edit Data Absensi Bulanan")
    df_to_edit = st.session_state.data_absensi.copy().reset_index(drop=True)
    
    edited_df = st.data_editor(df_to_edit, num_rows="dynamic", use_container_width=True, hide_index=True, key="edit_bulanan")
    
    if st.button("💾 Simpan Data Absensi", type="primary"):
        edited_df["No"] = range(1, len(edited_df) + 1)
        st.session_state.data_absensi = edited_df
        save_data(edited_df, FILE_ABSENSI)
        st.success("✅ Data Absensi Bulanan berhasil disimpan secara permanen!")

    # Kalkulasi
    df_calc = st.session_state.data_absensi.copy()
    for col in ["HBE", "S", "I", "A"]:
        df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)
    
    df_calc["JUMLAH"] = df_calc["S"] + df_calc["I"] + df_calc["A"]
    df_calc["KEHADIRAN %"] = (((df_calc["HBE"] - df_calc["JUMLAH"]) / df_calc["HBE"]) * 100).fillna(0).round(1)

    st.markdown("---")
    st.subheader("📊 Tabel Laporan Rekapitulasi")
    st.dataframe(df_calc, use_container_width=True, hide_index=True)

    # Download Excel Bulanan
    output_bulanan = io.BytesIO()
    with pd.ExcelWriter(output_bulanan, engine='openpyxl') as writer:
        df_calc.to_excel(writer, sheet_name="Rekap_Bulanan", index=False)
    
    st.download_button(
        label="📥 Download Laporan Absensi Bulanan (Excel)",
        data=output_bulanan.getvalue(),
        file_name=f"Rekap_Absensi_Bulanan_{kelas}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN HARIAN
# ------------------------------------------
with tab2:
    st.subheader("📅 Input & Catat Kehadiran Harian Siswa")
    
    with st.form("form_harian"):
        col_h1, col_h2, col_h3, col_h4, col_h5 = st.columns(5)
        tgl_harian = col_h1.date_input("Tanggal", date.today())
        total_siswa_aktif = len(st.session_state.data_absensi)
        
        jml_hadir = col_h2.number_input("Hadir", min_value=0, value=total_siswa_aktif)
        jml_sakit = col_h3.number_input("Sakit (S)", min_value=0, value=0)
        jml_izin = col_h4.number_input("Izin (I)", min_value=0, value=0)
        jml_alpha = col_h5.number_input("Alpha (A)", min_value=0, value=0)
        
        btn_simpan_harian = st.form_submit_button("➕ Tambah Rekap Harian")
        
        if btn_simpan_harian:
            tot_input = jml_hadir + jml_sakit + jml_izin + jml_alpha
            pct_hadir = round((jml_hadir / tot_input) * 100, 2) if tot_input > 0 else 0.0
            
            row_baru = {
                "Tanggal": str(tgl_harian),
                "Hadir": jml_hadir,
                "Sakit": jml_sakit,
                "Izin": jml_izin,
                "Alpha": jml_alpha,
                "Total Siswa": tot_input,
                "Persentase Kehadiran (%)": pct_hadir
            }
            
            df_harian_new = pd.concat([st.session_state.data_harian, pd.DataFrame([row_baru])], ignore_index=True)
            st.session_state.data_harian = df_harian_new
            save_data(df_harian_new, FILE_HARIAN)
            st.success(f"✅ Data kehadiran harian tanggal {tgl_harian} berhasil disimpan!")
            st.rerun()

    st.markdown("---")
    st.subheader("📈 Riwayat Persentase Kehadiran Harian")
    
    df_harian_edit = st.data_editor(st.session_state.data_harian, num_rows="dynamic", use_container_width=True, hide_index=True, key="edit_harian")
    
    if st.button("💾 Simpan Perubahan Data Harian"):
        st.session_state.data_harian = df_harian_edit
        save_data(df_harian_edit, FILE_HARIAN)
        st.success("✅ Perubahan data harian berhasil disimpan!")

    # Download Excel Harian
    output_harian = io.BytesIO()
    with pd.ExcelWriter(output_harian, engine='openpyxl') as writer:
        df_harian_edit.to_excel(writer, sheet_name="Kehadiran_Harian", index=False)
    
    st.download_button(
        label="📥 Download Laporan Kehadiran Per Hari (Excel)",
        data=output_harian.getvalue(),
        file_name=f"Laporan_Kehadiran_Harian_{kelas}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------------------------
# TAB 3: MUTASI SISWA (MASUK / KELUAR)
# ------------------------------------------
with tab3:
    st.subheader("🔄 Pencatatan Mutasi Siswa (Masuk / Keluar)")
    
    with st.form("form_mutasi"):
        col_m1, col_m2, col_m3 = st.columns(3)
        tgl_mutasi = col_m1.date_input("Tanggal Mutasi", date.today())
        nama_mutasi = col_m2.text_input("Nama Siswa")
        lp_mutasi = col_m3.selectbox("Jenis Kelamin", ["L", "P"])
        
        col_m4, col_m5, col_m6 = st.columns(3)
        nisn_mutasi = col_m4.text_input("NISN / Nomor Induk")
        jenis_mutasi = col_m5.selectbox("Jenis Mutasi", ["Masuk", "Keluar"])
        ket_mutasi = col_m6.text_input("Keterangan / Sekolah Asal atau Tujuan")
        
        btn_tambah_mutasi = st.form_submit_button("➕ Simpan Data Mutasi")
        
        if btn_tambah_mutasi:
            if nama_mutasi.strip() == "":
                st.error("Nama siswa mutasi wajib diisi!")
            else:
                row_mutasi = {
                    "Tanggal Mutasi": str(tgl_mutasi),
                    "Nama Siswa": nama_mutasi.upper().strip(),
                    "L/P": lp_mutasi,
                    "NISN": nisn_mutasi.strip(),
                    "Jenis Mutasi": jenis_mutasi,
                    "Keterangan / Sekolah Asal/Tujuan": ket_mutasi.strip()
                }
                
                df_mutasi_new = pd.concat([st.session_state.data_mutasi, pd.DataFrame([row_mutasi])], ignore_index=True)
                st.session_state.data_mutasi = df_mutasi_new
                save_data(df_mutasi_new, FILE_MUTASI)
                st.success(f"✅ Data mutasi siswa {nama_mutasi.upper()} berhasil ditambahkan!")
                st.rerun()

    st.markdown("---")
    st.subheader("📋 Daftar Riwayat Mutasi Siswa")
    
    df_mutasi_edit = st.data_editor(st.session_state.data_mutasi, num_rows="dynamic", use_container_width=True, hide_index=True, key="edit_mutasi")
    
    if st.button("💾 Simpan Perubahan Data Mutasi"):
        st.session_state.data_mutasi = df_mutasi_edit
        save_data(df_mutasi_edit, FILE_MUTASI)
        st.success("✅ Perubahan data mutasi berhasil disimpan!")

    # Download Excel Mutasi
    output_mutasi = io.BytesIO()
    with pd.ExcelWriter(output_mutasi, engine='openpyxl') as writer:
        df_mutasi_edit.to_excel(writer, sheet_name="Mutasi_Siswa", index=False)
    
    st.download_button(
        label="📥 Download Laporan Mutasi Siswa (Excel)",
        data=output_mutasi.getvalue(),
        file_name=f"Laporan_Mutasi_Siswa_{kelas}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
