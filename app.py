import streamlit as st
import pandas as pd
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import date

# ==========================================
# 0. KONFIGURASI PENYIMPANAN DATA PERMANEN
# ==========================================
FILE_ABSENSI = "database_absensi_siswa.csv"
FILE_HARIAN = "database_kehadiran_harian.csv"
FILE_MUTASI = "database_mutasi_siswa.csv"

# Inisialisasi Data Default Rekap Harian (Tanggal 1 - 31)
def generate_default_harian():
    rows = []
    # Contoh posisi hari minggu pada tanggal 5, 12, 19, 26 sesuai gambar
    minggu_days = [5, 12, 19, 26]
    for tgl in range(1, 32):
        if tgl in minggu_days:
            rows.append({
                "TGL": str(tgl),
                "Jumlah Siswa": "MINGGU",
                "S": "-",
                "I": "-",
                "A": "-",
                "Jumlah": "-",
                "Hadir %": "-",
                "Tidak hadir %": "-"
            })
        else:
            rows.append({
                "TGL": str(tgl),
                "Jumlah Siswa": 30,
                "S": 0,
                "I": 0,
                "A": 0,
                "Jumlah": 0,
                "Hadir %": "100%",
                "Tidak hadir %": "0%"
            })
    return pd.DataFrame(rows)

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
    DEFAULT_ABSENSI = pd.DataFrame([
        {"No": 1, "Nama Murid": "ABANG MUSHAWIR EDO", "L/P": "L", "Nomor Induk": "0012345678", "HBE": 25, "S": 0, "I": 0, "A": 0},
        {"No": 2, "Nama Murid": "AHMAD YANI", "L/P": "L", "Nomor Induk": "0012345679", "HBE": 25, "S": 0, "I": 0, "A": 2},
    ])
    st.session_state.data_absensi = load_data(FILE_ABSENSI, DEFAULT_ABSENSI)

if "data_harian" not in st.session_state:
    st.session_state.data_harian = load_data(FILE_HARIAN, generate_default_harian())

if "data_mutasi" not in st.session_state:
    DEFAULT_MUTASI = pd.DataFrame([
        {"Tanggal Mutasi": str(date.today()), "Nama Siswa": "BUDI SANTOSO", "L/P": "L", "NISN": "0019998877", "Jenis Mutasi": "Masuk", "Keterangan / Sekolah Asal/Tujuan": "SMPN 2 Nanga Mahap"}
    ])
    st.session_state.data_mutasi = load_data(FILE_MUTASI, DEFAULT_MUTASI)

# ==========================================
# 2. SIDEBAR INFORMASI SEKOLAH
# ==========================================
st.sidebar.header("🏫 Data Sekolah & Kelas")
nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SMP NEGERI 1 NANGA MAHAP")
tahun_pelajaran = st.sidebar.text_input("Tahun Pelajaran", "2025/2026")
kelas = st.sidebar.text_input("Kelas", "IX C")
bulan_tahun = st.sidebar.text_input("Bulan / Periode", "April 2026")

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
        st.success("✅ Data Absensi Bulanan berhasil disimpan!")

# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN PER HARI (FORMAT SESUAI GAMBAR)
# ------------------------------------------
with tab2:
    st.subheader("📅 Tabel Rekapitulasi Kehadiran Siswa Per Hari")
    st.caption("Isi nilai S, I, A atau atur kolom 'Jumlah Siswa' menjadi MINGGU untuk menandai hari libur.")

    df_harian_input = st.session_state.data_harian.copy()
    
    # Editor Tabel Interaktif
    edited_harian = st.data_editor(
        df_harian_input,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key="editor_harian_gambar"
    )

    if st.button("💾 Simpan & Hitung Ulang Persentase", type="primary"):
        # Melakukan perhitungan ulang sesuai aturan tabel Excel pada gambar
        processed_rows = []
        tot_s = 0
        tot_i = 0
        tot_a = 0
        tot_tidak_hadir = 0
        tot_keseluruhan_siswa = 0
        count_hari_efektif = 0

        for idx, row in edited_harian.iterrows():
            tgl_val = str(row["TGL"])
            js_val = str(row["Jumlah Siswa"]).strip().upper()

            if js_val == "MINGGU" or row["S"] == "-" or row["I"] == "-":
                processed_rows.append({
                    "TGL": tgl_val,
                    "Jumlah Siswa": "MINGGU",
                    "S": "-",
                    "I": "-",
                    "A": "-",
                    "Jumlah": "-",
                    "Hadir %": "-",
                    "Tidak hadir %": "-"
                })
            else:
                try:
                    jml_siswa = int(float(js_val))
                    s_val = int(float(row["S"]))
                    i_val = int(float(row["I"]))
                    a_val = int(float(row["A"]))
                except ValueError:
                    jml_siswa, s_val, i_val, a_val = 30, 0, 0, 0

                jml_th = s_val + i_val + a_val
                pct_th = round((jml_th / jml_siswa) * 100) if jml_siswa > 0 else 0
                pct_h = 100 - pct_th if jml_siswa > 0 else 100

                tot_s += s_val
                tot_i += i_val
                tot_a += a_val
                tot_tidak_hadir += jml_th
                tot_keseluruhan_siswa += jml_siswa
                count_hari_efektif += 1

                processed_rows.append({
                    "TGL": tgl_val,
                    "Jumlah Siswa": jml_siswa,
                    "S": s_val,
                    "I": i_val,
                    "A": a_val,
                    "Jumlah": jml_th,
                    "Hadir %": f"{pct_h}%",
                    "Tidak hadir %": f"{pct_th}%"
                })

        # Baris Total / JUMLAH Paling Bawah
        avg_pct_th = round((tot_tidak_hadir / tot_keseluruhan_siswa) * 100) if tot_keseluruhan_siswa > 0 else 0
        avg_pct_h = 100 - avg_pct_th if tot_keseluruhan_siswa > 0 else 100

        df_processed = pd.DataFrame(processed_rows)
        
        # Simpan ke Session State & File CSV
        st.session_state.data_harian = df_processed
        save_data(df_processed, FILE_HARIAN)
        st.success("✅ Perhitungan Rekap Harian Berhasil Diperbarui!")
        st.rerun()

    # Hitung Rekapitulasi Akhir untuk Ditampilkan di Baris Bawah
    df_current = st.session_state.data_harian.copy()
    tot_s, tot_i, tot_a, tot_th, tot_siswa = 0, 0, 0, 0, 0

    for idx, row in df_current.iterrows():
        if str(row["Jumlah Siswa"]).upper() != "MINGGU" and row["S"] != "-":
            try:
                tot_s += int(float(row["S"]))
                tot_i += int(float(row["I"]))
                tot_a += int(float(row["A"]))
                tot_th += int(float(row["Jumlah"]))
                tot_siswa += int(float(row["Jumlah Siswa"]))
            except ValueError:
                pass

    total_pct_th = round((tot_th / tot_siswa) * 100) if tot_siswa > 0 else 0
    total_pct_h = 100 - total_pct_th if tot_siswa > 0 else 100

    # Menampilkan Ringkasan Total Paling Bawah
    st.markdown("### 📊 Ringkasan Total Bulanan")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Sakit (S)", tot_s)
    c2.metric("Total Izin (I)", tot_i)
    c3.metric("Total Alpha (A)", tot_a)
    c4.metric("Total Tidak Hadir", tot_th)
    c5.metric("Rata-rata Hadir %", f"{total_pct_h}%")
    c6.metric("Rata-rata Tidak Hadir %", f"{total_pct_th}%")

    # Function Export ke Excel dengan Format & Warna Persis Gambar
    def generate_excel_harian(df_data, total_s, total_i, total_a, total_th, pct_h, pct_th):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Harian"

        # Format Border & Alignments
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal="center", vertical="center")
        bold_font = Font(bold=True)
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Header Multi-level
        ws.merge_cells("A1:A2")
        ws["A1"] = "TGL"
        ws.merge_cells("B1:B2")
        ws["B1"] = "Jumlah Siswa"
        
        ws.merge_cells("C1:E1")
        ws["C1"] = "Tidak hadir Karena"
        ws["C2"], ws["D2"], ws["E2"] = "S", "I", "A"

        ws.merge_cells("F1:F2")
        ws["F1"] = "Jumlah"

        ws.merge_cells("G1:H1")
        ws["G1"] = "Presentase"
        ws["G2"], ws["H2"] = "Hadir %", "Tidak hadir %"

        # Apply header styling
        for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=8):
            for cell in row:
                cell.alignment = center_align
                cell.font = bold_font
                cell.border = thin_border

        # Isi Data Baris 1-31
        curr_row = 3
        for idx, row in df_data.iterrows():
            if str(row["Jumlah Siswa"]).upper() == "MINGGU":
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=8)
                ws.cell(row=curr_row, column=1, value=int(row["TGL"])).alignment = center_align
                cell_m = ws.cell(row=curr_row, column=2, value="MINGGU")
                cell_m.alignment = center_align
                cell_m.font = bold_font
                
                for col in range(1, 9):
                    c = ws.cell(row=curr_row, column=col)
                    c.fill = red_fill
                    c.border = thin_border
            else:
                ws.cell(row=curr_row, column=1, value=int(row["TGL"])).alignment = center_align
                ws.cell(row=curr_row, column=2, value=int(row["Jumlah Siswa"])).alignment = center_align
                ws.cell(row=curr_row, column=3, value=int(row["S"])).alignment = center_align
                ws.cell(row=curr_row, column=4, value=int(row["I"])).alignment = center_align
                ws.cell(row=curr_row, column=5, value=int(row["A"])).alignment = center_align
                ws.cell(row=curr_row, column=6, value=int(row["Jumlah"])).alignment = center_align
                ws.cell(row=curr_row, column=7, value=str(row["Hadir %"])).alignment = center_align
                ws.cell(row=curr_row, column=8, value=str(row["Tidak hadir %"])).alignment = center_align

                for col in range(1, 9):
                    ws.cell(row=curr_row, column=col).border = thin_border

            curr_row += 1

        # Baris JUMLAH Paling Bawah
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=2)
        ws.cell(row=curr_row, column=1, value="JUMLAH").alignment = center_align
        ws.cell(row=curr_row, column=3, value=total_s).alignment = center_align
        ws.cell(row=curr_row, column=4, value=total_i).alignment = center_align
        ws.cell(row=curr_row, column=5, value=total_a).alignment = center_align
        ws.cell(row=curr_row, column=6, value=total_th).alignment = center_align
        
        c_h = ws.cell(row=curr_row, column=7, value=f"{pct_h}%")
        c_th = ws.cell(row=curr_row, column=8, value=f"{pct_th}%")
        c_h.alignment = center_align
        c_th.alignment = center_align
        c_h.fill = yellow_fill
        c_th.fill = yellow_fill

        for col in range(1, 9):
            cell = ws.cell(row=curr_row, column=col)
            cell.font = bold_font
            cell.border = thin_border

        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    # Tombol Download Excel Format Gambar
    excel_bytes = generate_excel_harian(df_current, tot_s, tot_i, tot_a, tot_th, total_pct_h, total_pct_th)
    st.download_button(
        label="📥 Download Laporan Rekap Harian (Format Excel Seperti Gambar)",
        data=excel_bytes,
        file_name=f"Rekap_Kehadiran_Harian_{kelas}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ------------------------------------------
# TAB 3: MUTASI SISWA
# ------------------------------------------
with tab3:
    st.subheader("🔄 Pencatatan Mutasi Siswa (Masuk / Keluar)")
    df_mutasi_edit = st.data_editor(st.session_state.data_mutasi, num_rows="dynamic", use_container_width=True, hide_index=True, key="edit_mutasi")
    
    if st.button("💾 Simpan Data Mutasi"):
        st.session_state.data_mutasi = df_mutasi_edit
        save_data(df_mutasi_edit, FILE_MUTASI)
        st.success("✅ Perubahan data mutasi berhasil disimpan!")
