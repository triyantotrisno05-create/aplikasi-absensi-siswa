import streamlit as st
import pandas as pd
from datetime import date

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="Aplikasi Absensi Siswa",
    page_icon="📝",
    layout="wide"
)

# ==========================================
# 2. INISIALISASI SESSION STATE
# ==========================================
# Menginisialisasi data absensi awal jika belum ada di session_state
if "df_absensi" not in st.session_state:
    data_awal = {
        "NIS": ["1001", "1002", "1003", "1004", "1005"],
        "Nama Siswa": ["Ahmad Fauzi", "Budi Santoso", "Citra Dewi", "Deni Pratama", "Eka Putri"],
        "Kelas": ["10 A", "10 A", "10 A", "10 A", "10 A"],
        "Tanggal": [str(date.today())] * 5,
        "Status": ["Hadir", "Hadir", "Izin", "Sakit", "Hadir"],
        "Catatan": ["-", "-", "Acara Keluarga", "Demam", "-"]
    }
    st.session_state.df_absensi = pd.DataFrame(data_awal)

if "data_version" not in st.session_state:
    st.session_state.data_version = 0

# ==========================================
# 3. JUDUL & HEADER
# ==========================================
st.title("📝 Input / Edit Data Absensi Siswa")
st.write("Silakan kelola dan perbarui data absensi siswa di bawah ini.")

st.divider()

# ==========================================
# 4. FILTER DATA (OPSIONAL)
# ==========================================
col_filter1, col_filter2 = st.columns(2)

with col_filter1:
    pilihan_kelas = st.selectbox(
        "Filter Kelas:",
        options=["Semua"] + list(st.session_state.df_absensi["Kelas"].unique())
    )

with col_filter2:
    pilihan_status = st.selectbox(
        "Filter Status:",
        options=["Semua", "Hadir", "Izin", "Sakit", "Alpha"]
    )

# Filter dataframe berdasarkan pilihan
df_to_edit = st.session_state.df_absensi.copy()

if pilihan_kelas != "Semua":
    df_to_edit = df_to_edit[df_to_edit["Kelas"] == pilihan_kelas]

if pilihan_status != "Semua":
    df_to_edit = df_to_edit[df_to_edit["Status"] == pilihan_status]

# PENTING: Reset index agar tidak ada error index terduplikasi pada data_editor
df_to_edit = df_to_edit.reset_index(drop=True)

# ==========================================
# 5. DATA EDITOR (PERBAIKAN STREAMLIT ERROR)
# ==========================================
st.subheader("Tabel Data Absensi")

# Menggunakan key statis dan pengurusan tipe data yang aman
edited_df = st.data_editor(
    df_to_edit,
    column_config={
        "NIS": st.column_config.TextColumn("NIS", disabled=True),
        "Nama Siswa": st.column_config.TextColumn("Nama Siswa", required=True),
        "Kelas": st.column_config.TextColumn("Kelas"),
        "Tanggal": st.column_config.DateColumn("Tanggal", format="YYYY-MM-DD"),
        "Status": st.column_config.SelectboxColumn(
            "Status Absensi",
            options=["Hadir", "Izin", "Sakit", "Alpha"],
            required=True
        ),
        "Catatan": st.column_config.TextColumn("Catatan / Alasan")
    },
    num_rows="dynamic",  # Memungkinkan tambah / hapus baris
    use_container_width=True,
    hide_index=True,
    key="editor_absensi_siswa"  # Key statis yang stabil
)

# ==========================================
# 6. TOMBOL SIMPAN & ACTION
# ==========================================
col_btn1, col_btn2 = st.columns([1, 4])

with col_btn1:
    if st.button("💾 Simpan Perubahan", type="primary", use_container_width=True):
        # Update data di session state
        st.session_state.df_absensi = edited_df.copy()
        st.session_state.data_version += 1
        st.success("Data absensi berhasil diperbarui!")

with col_btn2:
    if st.button("🔄 Reset Data Ke Default", use_container_width=True):
        st.session_state.clear()
        st.rerun()

# ==========================================
# 7. RINGKASAN DATA (METRICS)
# ==========================================
st.divider()
st.subheader("📊 Ringkasan Absensi Hari Ini")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Hadir", len(st.session_state.df_absensi[st.session_state.df_absensi["Status"] == "Hadir"]))
m2.metric("Total Izin", len(st.session_state.df_absensi[st.session_state.df_absensi["Status"] == "Izin"]))
m3.metric("Total Sakit", len(st.session_state.df_absensi[st.session_state.df_absensi["Status"] == "Sakit"]))
m4.metric("Total Alpha", len(st.session_state.df_absensi[st.session_state.df_absensi["Status"] == "Alpha"]))
