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

def load_data_harian(filepath):
    default_df = generate_default_harian()
    if os.path.exists(filepath):
        try:
            df = pd.read_csv(filepath)
            required_cols = ["TGL", "Jumlah Siswa", "S", "I", "A", "Jumlah", "Hadir %", "Tidak hadir %"]
            if not all(col in df.columns for col in required_cols):
                default_df.to_csv(filepath, index=False)
                return default_df.copy()
            return df
        except Exception:
            return default_df.copy()
    return default_df.copy()

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
        {"No": 3, "Nama Murid": "AL JAMI", "L/P": "L", "Nomor Induk": "0012345680", "HBE": 25, "S": 1, "I": 0, "A": 3},
        {"No": 4, "Nama Murid": "AYU NINGSIH", "L/P": "P", "Nomor Induk": "0012345681", "HBE": 25, "S": 0, "I": 0, "A": 0},
        {"No": 5, "Nama Murid": "EPRI SASKIA", "L/P": "P", "Nomor Induk": "0012345682", "HBE": 25, "S": 0, "I": 0, "A": 2},
    ])
    st.session_state.data_absensi = load_data(FILE_ABSENSI, DEFAULT_ABSENSI)

if "data_harian" not in st.session_state:
    st.session_state.data_harian = load_data_harian(FILE_HARIAN)

if "data_mutasi" not in st.session_state:
    DEFAULT_MUTASI = pd.DataFrame([
        {"Tanggal Mutasi": str(date.today()), "Nama Siswa": "BUDI SANTOSO", "L/P": "L", "NISN": "0019998877", "Jenis Mutasi": "Masuk", "Keterangan / Sekolah Asal/Tujuan": "SMPN 2 Nanga Mahap"}
    ])
    st.session_state.data_mutasi = load_data(FILE_MUTASI, DEFAULT_MUTASI)

# ==========================================
# 2. SIDEBAR INFORMASI SEKOLAH & TTD
# ==========================================
st.sidebar.header("🏫 Data Sekolah & Kelas")
nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SMP NEGERI 1 NANGA MAHAP")
tahun_pelajaran = st.sidebar.text_input("Tahun Pelajaran", "2025/2026")
kelas = st.sidebar.text_input("Kelas", "IX C")
bulan_tahun = st.sidebar.text_input("Bulan / Periode", "April 2026")

st.sidebar.divider()
st.sidebar.header("✍️ Data Tanda Tangan")
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
    st.subheader("📝 Edit Data Absensi Bulanan Siswa")
    df_to_edit = st.session_state.data_absensi.copy().reset_index(drop=True)
    edited_df = st.data_editor(df_to_edit, num_rows="dynamic", use_container_width=True, hide_index=True, key="edit_bulanan")
    
    if st.button("💾 Simpan Data Absensi", type="primary"):
        edited_df["No"] = range(1, len(edited_df) + 1)
        st.session_state.data_absensi = edited_df
        save_data(edited_df, FILE_ABSENSI)
        st.success("✅ Data Absensi Bulanan berhasil disimpan!")

    # Kalkulasi Otomatis
    df_calc = st.session_state.data_absensi.copy()
    for col in ["HBE", "S", "I", "A"]:
        df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)
    
    df_calc["JUMLAH"] = df_calc["S"] + df_calc["I"] + df_calc["A"]
    df_calc["KEHADIRAN %"] = (((df_calc["HBE"] - df_calc["JUMLAH"]) / df_calc["HBE"]) * 100).fillna(0).round(1)

    st.markdown("---")
    st.subheader("📊 Tabel Laporan Rekapitulasi Bulanan")
    st.dataframe(df_calc, use_container_width=True, hide_index=True)

    # Function Generate Excel Siap Cetak untuk Bulanan
    def generate_excel_bulanan_siap_cetak(df, sekolah, kls, th_ajar, bulan, ktt, tgl_c, wali, nip):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Absensi Bulanan"

        # Styles
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        bold_font = Font(name="Arial", size=10, bold=True)
        normal_font = Font(name="Arial", size=10)
        title_font = Font(name="Arial", size=12, bold=True)
        gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

        # 1. KOP / JUDUL LAPORAN
        ws.merge_cells("A1:J1")
        ws["A1"] = f"LAPORAN REKAPITULASI ABSENSI SISWA BULANAN"
        ws["A1"].font = title_font
        ws["A1"].alignment = center_align

        ws.merge_cells("A2:J2")
        ws["A2"] = sekolah.upper()
        ws["A2"].font = title_font
        ws["A2"].alignment = center_align

        ws["A4"] = f"Kelas: {kls}"
        ws["A4"].font = bold_font
        ws["F4"] = f"Bulan / Periode: {bulan}"
        ws["F4"].font = bold_font

        ws["A5"] = f"Tahun Pelajaran: {th_ajar}"
        ws["A5"].font = bold_font

        # 2. HEADER TABEL
        headers = ["No", "Nama Murid", "L/P", "Nomor Induk", "HBE", "Sakit (S)", "Izin (I)", "Alpha (A)", "Jumlah", "Kehadiran %"]
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=7, column=col_idx, value=header)
            cell.font = bold_font
            cell.alignment = center_align
            cell.border = thin_border
            cell.fill = gray_fill

        # 3. ISI DATA
        current_row = 8
        for _, row in df.iterrows():
            ws.cell(row=current_row, column=1, value=int(row.get("No", 0))).alignment = center_align
            ws.cell(row=current_row, column=2, value=str(row.get("Nama Murid", ""))).alignment = left_align
            ws.cell(row=current_row, column=3, value=str(row.get("L/P", ""))).alignment = center_align
            ws.cell(row=current_row, column=4, value=str(row.get("Nomor Induk", ""))).alignment = center_align
            ws.cell(row=current_row, column=5, value=int(row.get("HBE", 0))).alignment = center_align
            ws.cell(row=current_row, column=6, value=int(row.get("S", 0))).alignment = center_align
            ws.cell(row=current_row, column=7, value=int(row.get("I", 0))).alignment = center_align
            ws.cell(row=current_row, column=8, value=int(row.get("A", 0))).alignment = center_align
            ws.cell(row=current_row, column=9, value=int(row.get("JUMLAH", 0))).alignment = center_align
            ws.cell(row=current_row, column=10, value=f"{row.get('KEHADIRAN %', 0)}%").alignment = center_align

            for col in range(1, 11):
                ws.cell(row=current_row, column=col).border = thin_border
                ws.cell(row=current_row, column=col).font = normal_font

            current_row += 1

        # 4. BARIS TOTAL / REKAP
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
        ws.cell(row=current_row, column=1, value="TOTAL / RATA-RATA KELAS").alignment = center_align
        
        ws.cell(row=current_row, column=5, value=df["HBE"].sum()).alignment = center_align
        ws.cell(row=current_row, column=6, value=df["S"].sum()).alignment = center_align
        ws.cell(row=current_row, column=7, value=df["I"].sum()).alignment = center_align
        ws.cell(row=current_row, column=8, value=df["A"].sum()).alignment = center_align
        ws.cell(row=current_row, column=9, value=df["JUMLAH"].sum()).alignment = center_align
        ws.cell(row=current_row, column=10, value=f"{round(df['KEHADIRAN %'].mean(), 1)}%").alignment = center_align

        for col in range(1, 11):
            cell = ws.cell(row=current_row, column=col)
            cell.font = bold_font
            cell.border = thin_border
            cell.fill = gray_fill

        # 5. TANDA TANGAN WALI KELAS
        current_row += 3
        tgl_str = tgl_c.strftime("%d %B %Y") if isinstance(tgl_c, date) else str(tgl_c)
        ws.cell(row=current_row, column=7, value=f"{ktt}, {tgl_str}").font = normal_font
        current_row += 1
        ws.cell(row=current_row, column=7, value="Wali Kelas,").font = normal_font
        current_row += 4
        
        c_wali = ws.cell(row=current_row, column=7, value=wali)
        c_wali.font = bold_font
        current_row += 1
        ws.cell(row=current_row, column=7, value=f"NIP. {nip}").font = normal_font

        # Auto Column Width
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    # Tombol Download Excel Siap Cetak
    excel_bulanan_bytes = generate_excel_bulanan_siap_cetak(
        df_calc, nama_sekolah, kelas, tahun_pelajaran, bulan_tahun, tempat_cetak, tgl_cetak, wali_kelas, nip_wali
    )

    st.download_button(
        label="🖨️ Download Laporan Rekapitulasi Bulanan Siap Cetak (Excel)",
        data=excel_bulanan_bytes,
        file_name=f"Laporan_Absensi_Bulanan_{kelas}_{bulan_tahun}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN PER HARI
# ------------------------------------------
with tab2:
    st.subheader("📅 Tabel Rekapitulasi Kehadiran Siswa Per Hari")
    st.caption("Isi nilai S, I, A atau ketik 'MINGGU' pada kolom Jumlah Siswa untuk menandai hari libur.")

    df_harian_input = st.session_state.data_harian.copy()
    
    edited_harian = st.data_editor(
        df_harian_input,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key="editor_harian_v2"
    )

    if st.button("💾 Simpan & Hitung Ulang Persentase", type="primary"):
        processed_rows = []
        for idx, row in edited_harian.iterrows():
            tgl_val = str(row.get("TGL", idx + 1))
            js_val = str(row.get("Jumlah Siswa", "30")).strip().upper()

            if js_val == "MINGGU" or str(row.get("S", "")).strip() == "-":
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
                    s_val = int(float(row.get("S", 0)))
                    i_val = int(float(row.get("I", 0)))
                    a_val = int(float(row.get("A", 0)))
                except (ValueError, TypeError):
                    jml_siswa, s_val, i_val, a_val = 30, 0, 0, 0

                jml_th = s_val + i_val + a_val
                pct_th = round((jml_th / jml_siswa) * 100) if jml_siswa > 0 else 0
                pct_h = 100 - pct_th if jml_siswa > 0 else 100

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

        df_processed = pd.DataFrame(processed_rows)
        st.session_state.data_harian = df_processed
        save_data(df_processed, FILE_HARIAN)
        st.success("✅ Perhitungan Rekap Harian Berhasil Diperbarui!")
        st.rerun()

    # Hitung Akumulasi Total
    df_current = st.session_state.data_harian.copy()
    tot_s, tot_i, tot_a, tot_th, tot_siswa = 0, 0, 0, 0, 0

    if "Jumlah Siswa" in df_current.columns and "S" in df_current.columns:
        for idx, row in df_current.iterrows():
            js_str = str(row["Jumlah Siswa"]).strip().upper()
            if js_str != "MINGGU" and str(row["S"]).strip() != "-":
                try:
                    tot_s += int(float(row["S"]))
                    tot_i += int(float(row["I"]))
                    tot_a += int(float(row["A"]))
                    tot_th += int(float(row["Jumlah"]))
                    tot_siswa += int(float(row["Jumlah Siswa"]))
                except (ValueError, TypeError):
                    pass

    total_pct_th = round((tot_th / tot_siswa) * 100) if tot_siswa > 0 else 0
    total_pct_h = 100 - total_pct_th if tot_siswa > 0 else 100

    st.markdown("---")
    st.markdown("### 📊 Ringkasan Total Bulanan")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Sakit (S)", tot_s)
    c2.metric("Total Izin (I)", tot_i)
    c3.metric("Total Alpha (A)", tot_a)
    c4.metric("Total Tidak Hadir", tot_th)
    c5.metric("Rata-rata Hadir %", f"{total_pct_h}%")
    c6.metric("Rata-rata Tidak Hadir %", f"{total_pct_th}%")

    def generate_excel_harian(df_data, total_s, total_i, total_a, total_th, pct_h, pct_th):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Harian"

        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal="center", vertical="center")
        bold_font = Font(bold=True)
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

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

        for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=8):
            for cell in row:
                cell.alignment = center_align
                cell.font = bold_font
                cell.border = thin_border

        curr_row = 3
        for idx, row in df_data.iterrows():
            if str(row.get("Jumlah Siswa", "")).upper() == "MINGGU":
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=8)
                ws.cell(row=curr_row, column=1, value=str(row.get("TGL", idx + 1))).alignment = center_align
                cell_m = ws.cell(row=curr_row, column=2, value="MINGGU")
                cell_m.alignment = center_align
                cell_m.font = bold_font
                
                for col in range(1, 9):
                    c = ws.cell(row=curr_row, column=col)
                    c.fill = red_fill
                    c.border = thin_border
            else:
                ws.cell(row=curr_row, column=1, value=str(row.get("TGL", idx + 1))).alignment = center_align
                ws.cell(row=curr_row, column=2, value=str(row.get("Jumlah Siswa", 30))).alignment = center_align
                ws.cell(row=curr_row, column=3, value=str(row.get("S", 0))).alignment = center_align
                ws.cell(row=curr_row, column=4, value=str(row.get("I", 0))).alignment = center_align
                ws.cell(row=curr_row, column=5, value=str(row.get("A", 0))).alignment = center_align
                ws.cell(row=curr_row, column=6, value=str(row.get("Jumlah", 0))).alignment = center_align
                ws.cell(row=curr_row, column=7, value=str(row.get("Hadir %", "100%"))).alignment = center_align
                ws.cell(row=curr_row, column=8, value=str(row.get("Tidak hadir %", "0%"))).alignment = center_align

                for col in range(1, 9):
                    ws.cell(row=curr_row, column=col).border = thin_border

            curr_row += 1

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

    excel_bytes = generate_excel_harian(df_current, tot_s, tot_i, tot_a, tot_th, total_pct_h, total_pct_th)
    st.download_button(
        label="📥 Download Laporan Rekap Harian (Excel Sesuai Format Gambar)",
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
