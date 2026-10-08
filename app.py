import streamlit as st
import pandas as pd
from datetime import date

# 1. Konfigurasi Halaman
st.set_page_config(
    page_title="Aplikasi Absensi Siswa",
    page_icon="📝",
    layout="wide"
)

st.title("📝 Input / Edit Data Absensi Siswa")
st.write("Kelola dan perbarui data absensi siswa di bawah ini.")
st.divider()

# 2. Inisialisasi Session State Data Awal
if "df_absensi" not in st.session_state:
    st.session_state.df_absensi = pd.DataFrame([
        {"NIS": "1001", "Nama Siswa": "Ahmad Fauzi", "Kelas": "10 A", "Tanggal": date.today(), "Status": "Hadir", "Catatan": "-"},
        {"NIS": "1002", "Nama Siswa": "Budi Santoso", "Kelas": "10 A", "Tanggal": date.today(), "Status": "Hadir", "Catatan": "-"},
        {"NIS": "1003", "Nama Siswa": "Citra Dewi", "Kelas": "10 A", "Tanggal": date.today(), "Status": "Izin", "Catatan": "Acara Keluarga"},
        {"NIS": "1004", "Nama Siswa": "Deni Pratama", "Kelas": "10 A", "Tanggal": date.today(), "Status": "Sakit", "Catatan": "Demam"},
        {"NIS": "1005", "Nama Siswa": "Eka Putri", "Kelas": "10 A", "Tanggal": date.today(), "Status": "Hadir", "Catatan": "-"},
    ])

# 3. Filter Data
col1, col2 = st.columns(2)
with col1:
    list_kelas = ["Semua"] + list(st.session_state.df_absensi["Kelas"].astype(str).unique())
    pilihan_kelas = st.selectbox("Filter Kelas:", list_kelas)

with col2:
    pilihan_status = st.selectbox("Filter Status:", ["Semua", "Hadir", "Izin", "Sakit", "Alpha"])

# Salin data untuk diedit
df_to_edit = st.session_state.df_absensi.copy()

if pilihan_kelas != "Semua":
    df_to_edit = df_to_edit[df_to_edit["Kelas"] == pilihan_kelas]

if pilihan_status != "Semua":
    df_to_edit = df_to_edit[df_to_edit["Status"] == pilihan_status]

# PENTING: Reset indeks agar Streamlit tidak bingung
df_to_edit = df_to_edit.reset_index(drop=True)

# 4. Data Editor Sederhana (Tanpa column_config rumit yang memicu crash)
edited_df = st.data_editor(
    df_to_edit,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="tabel_absensi_v1"
)

# 5. Tombol Aksi
st.write("")
col_btn1, col_btn2 = st.columns([1, 4])

with col_btn1:
    if st.button("💾 Simpan Perubahan", type="primary", use_container_width=True):
        st.session_state.df_absensi = edited_df.copy()
        st.success("✅ Data absensi berhasil disimpan!")
        st.rerun()

with col_btn2:
    if st.button("🔄 Reset Data Default", use_container_width=True):
        del st.session_state["df_absensi"]
        st.rerun()

# 6. Ringkasan
st.divider()
st.subheader("📊 Ringkasan Absensi")

df_curr = st.session_state.df_absensi
m1, m2, m3, m4 = st.columns(4)
m1.metric("Hadir", len(df_curr[df_curr["Status"] == "Hadir"]))
m2.metric("Izin", len(df_curr[df_curr["Status"] == "Izin"]))
m3.metric("Sakit", len(df_curr[df_curr["Status"] == "Sakit"]))
m4.metric("Alpha", len(df_curr[df_curr["Status"] == "Alpha"]))
