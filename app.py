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

def generate_default_harian():
    rows = []
    minggu_days = [5, 12, 19, 26]
    for tgl in range(1, 32):
        if tgl in minggu_days:
            rows.append({
                "TGL": str(tgl),
                "Jumlah Siswa": "MINGGU",
                "S": "-", "I": "-", "A": "-", "Jumlah": "-",
                "Hadir %": "-", "Tidak hadir %": "-"
            })
        else:
            rows.append({
                "TGL": str(tgl),
                "Jumlah Siswa": 30,
                "S": 0, "I": 0, "A": 0, "Jumlah": 0,
                "Hadir %": "100%", "Tidak hadir %": "0%"
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
bulan_tahun = st.sidebar.text_input("Bulan / Periode", "30 April 2026")

st.sidebar.divider()
st.sidebar.header("✍️ Data Tanda Tangan")
tempat_cetak = st.sidebar.text_input("Kota / Tempat Laporan", "Nanga Mahap")
tgl_cetak = st.sidebar.text_input("Tanggal Cetak TTD", "30 April 2026")
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "Triyanto trisno.S.Pd")
nip_wali = st.sidebar.text_input("NIP Wali Kelas", "1993052020242110006")

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
    
    df_calc = st.session_state.data_absensi.copy()
    for col in ["HBE", "S", "I", "A"]:
        df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)
    
    df_calc["JUMLAH"] = df_calc["S"] + df_calc["I"] + df_calc["A"]
    df_calc["JUMLAH HADIR"] = df_calc["HBE"] - df_calc["JUMLAH"]
    
    df_calc["S %"] = df_calc.apply(lambda r: round((r["S"] / r["HBE"]) * 100) if r["HBE"] > 0 else 0, axis=1)
    df_calc["I %"] = df_calc.apply(lambda r: round((r["I"] / r["HBE"]) * 100) if r["HBE"] > 0 else 0, axis=1)
    df_calc["A %"] = df_calc.apply(lambda r: round((r["A"] / r["HBE"]) * 100) if r["HBE"] > 0 else 0, axis=1)
    df_calc["PRESENTASE KEHADIRAN"] = df_calc.apply(lambda r: round((r["JUMLAH HADIR"] / r["HBE"]) * 100) if r["HBE"] > 0 else 100, axis=1)

    edited_df = st.data_editor(
        df_calc[["No", "Nama Murid", "L/P", "Nomor Induk", "HBE", "S", "I", "A", "JUMLAH", "JUMLAH HADIR", "PRESENTASE KEHADIRAN"]],
        disabled=["No", "JUMLAH", "JUMLAH HADIR", "PRESENTASE KEHADIRAN"],
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="edit_bulanan_v3"
    )
    
    if st.button("💾 Simpan Data Absensi", type="primary"):
        save_df = edited_df[["No", "Nama Murid", "L/P", "Nomor Induk", "HBE", "S", "I", "A"]].copy()
        save_df["No"] = range(1, len(save_df) + 1)
        st.session_state.data_absensi = save_df
        save_data(save_df, FILE_ABSENSI)
        st.success("✅ Data Absensi Bulanan berhasil disimpan!")
        st.rerun()

    def generate_excel_persis_gambar(df, tempat_c, tgl_c, wali, nip):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekapitulasi Bulanan"

        font_bold = Font(name="Calibri", size=11, bold=True)
        font_normal = Font(name="Calibri", size=11)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left = Alignment(horizontal="left", vertical="center")
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

        fill_green = PatternFill(start_color="66FF66", end_color="66FF66", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        ws.merge_cells("A1:A2"); ws["A1"] = "NO"
        ws.merge_cells("B1:B2"); ws["B1"] = "NAMA MURID"
        ws.merge_cells("C1:C2"); ws["C1"] = "L/P"
        ws.merge_cells("D1:D2"); ws["D1"] = "NOMOR INDUK"
        ws.merge_cells("E1:E2"); ws["E1"] = "HBE"
        
        ws.merge_cells("F1:H1"); ws["F1"] = "ABSENSI"
        ws["F2"] = "S"; ws["G2"] = "I"; ws["H2"] = "A"

        ws.merge_cells("I1:I2"); ws["I1"] = "JUMLAH"
        ws.merge_cells("J1:J2"); ws["J1"] = "JUMLAH HADIR"

        ws.merge_cells("K1:M1"); ws["K1"] = "PRESENTASE"
        ws["K2"] = "S"; ws["L2"] = "I"; ws["M2"] = "A"

        ws.merge_cells("N1:N2"); ws["N1"] = "PRESENTASE KEHADIRAN"

        for r in range(1, 3):
            for c in range(1, 15):
                cell = ws.cell(row=r, column=c)
                cell.font = font_bold
                cell.alignment = align_center
                cell.border = thin_border

        curr_row = 3
        count_l = 0
        count_p = 0

        for idx, row in df.iterrows():
            lp_val = str(row.get("L/P", "")).strip().upper()
            if lp_val == "L": count_l += 1
            elif lp_val == "P": count_p += 1

            ws.cell(row=curr_row, column=1, value=int(row.get("No", idx+1))).alignment = align_center
            ws.cell(row=curr_row, column=2, value=str(row.get("Nama Murid", ""))).alignment = align_left
            
            c_lp = ws.cell(row=curr_row, column=3, value=lp_val)
            c_lp.alignment = align_center
            c_lp.fill = fill_green

            c_nis = ws.cell(row=curr_row, column=4, value=str(row.get("Nomor Induk", "")))
            c_nis.alignment = align_center
            c_nis.fill = fill_green

            ws.cell(row=curr_row, column=5, value=int(row.get("HBE", 25))).alignment = align_center
            ws.cell(row=curr_row, column=6, value=int(row.get("S", 0))).alignment = align_center
            ws.cell(row=curr_row, column=7, value=int(row.get("I", 0))).alignment = align_center
            ws.cell(row=curr_row, column=8, value=int(row.get("A", 0))).alignment = align_center
            ws.cell(row=curr_row, column=9, value=int(row.get("JUMLAH", 0))).alignment = align_center
            ws.cell(row=curr_row, column=10, value=int(row.get("JUMLAH HADIR", 0))).alignment = align_center
            
            ws.cell(row=curr_row, column=11, value=f"{int(row.get('S %', 0))}%").alignment = align_center
            ws.cell(row=curr_row, column=12, value=f"{int(row.get('I %', 0))}%").alignment = align_center
            ws.cell(row=curr_row, column=13, value=f"{int(row.get('A %', 0))}%").alignment = align_center
            ws.cell(row=curr_row, column=14, value=f"{int(row.get('PRESENTASE KEHADIRAN', 100))}%").alignment = align_center

            for c in range(1, 15):
                cell = ws.cell(row=curr_row, column=c)
                cell.font = font_bold if c in [1, 3, 5, 6, 7, 8, 9, 10, 14] else font_normal
                cell.border = thin_border

            curr_row += 1

        tot_sakit = df["S"].sum()
        tot_izin = df["I"].sum()
        tot_alpha = df["A"].sum()
        tot_absen = df["JUMLAH"].sum()
        tot_hbe = df["HBE"].sum()
        tot_hadir = df["JUMLAH HADIR"].sum()

        avg_s_pct = round((tot_sakit / tot_hbe)*100) if tot_hbe > 0 else 0
        avg_i_pct = round((tot_izin / tot_hbe)*100) if tot_hbe > 0 else 0
        avg_a_pct = round((tot_alpha / tot_hbe)*100) if tot_hbe > 0 else 0
        avg_kehadiran = round((tot_hadir / tot_hbe)*100) if tot_hbe > 0 else 100

        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=2)
        ws.cell(row=curr_row, column=1, value="JUMLAH").alignment = align_center
        
        c_tot_siswa = ws.cell(row=curr_row, column=3, value=len(df))
        c_tot_siswa.alignment = align_center

        ws.cell(row=curr_row, column=6, value=tot_sakit).alignment = align_center
        ws.cell(row=curr_row, column=7, value=tot_izin).alignment = align_center
        ws.cell(row=curr_row, column=8, value=tot_alpha).alignment = align_center
        ws.cell(row=curr_row, column=9, value=tot_absen).alignment = align_center
        ws.cell(row=curr_row, column=10, value=tot_hadir).alignment = align_center
        
        ws.cell(row=curr_row, column=11, value=f"{avg_s_pct}%").alignment = align_center
        ws.cell(row=curr_row, column=12, value=f"{avg_i_pct}%").alignment = align_center
        ws.cell(row=curr_row, column=13, value=f"{avg_a_pct}%").alignment = align_center
        ws.cell(row=curr_row, column=14, value=f"{avg_kehadiran}%").alignment = align_center

        for c in range(1, 15):
            cell = ws.cell(row=curr_row, column=c)
            cell.font = font_bold
            cell.border = thin_border
            cell.fill = fill_yellow

        curr_row += 2
        ws.cell(row=curr_row, column=1, value="Laki - Laki").font = font_bold
        ws.cell(row=curr_row, column=3, value=":").font = font_bold
        ws.cell(row=curr_row, column=4, value=count_l).font = font_bold

        curr_row += 1
        ws.cell(row=curr_row, column=1, value="Perempuan").font = font_bold
        ws.cell(row=curr_row, column=3, value=":").font = font_bold
        ws.cell(row=curr_row, column=4, value=count_p).font = font_bold

        curr_row += 1
        ws.cell(row=curr_row, column=1, value="Jumlah akhir bulan").font = font_bold
        ws.cell(row=curr_row, column=3, value=":").font = font_bold
        ws.cell(row=curr_row, column=4, value=count_l + count_p).font = font_bold

        ws.cell(row=curr_row-2, column=10, value=f"{tempat_c}, {tgl_c}").font = font_bold
        ws.cell(row=curr_row-1, column=10, value=f"Wali Kelas  {kelas}").font = font_bold

        ws.cell(row=curr_row+3, column=10, value=wali).font = Font(name="Calibri", size=11, bold=True, underline="single")
        ws.cell(row=curr_row+4, column=10, value=f"NIP. {nip}").font = font_bold

        ws.column_dimensions['A'].width = 6
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 6
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 8
        for col_letter in ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
            ws.column_dimensions[col_letter].width = 12

        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    excel_bulanan_bytes = generate_excel_persis_gambar(
        df_calc, tempat_cetak, tgl_cetak, wali_kelas, nip_wali
    )

    st.download_button(
        label="📥 UNDUH REKAPITULASI BULANAN (EXCEL)",
        data=excel_bulanan_bytes,
        file_name=f"Rekapitulasi_Bulanan_{kelas}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN PER HARI (FORMAT PERSIS GAMBAR)
# ------------------------------------------
with tab2:
    st.subheader("📅 Persentase Kehadiran Per Hari (Format Persis Gambar Upload)")
    st.caption("Isi nilai S, I, A atau tulis 'MINGGU' pada kolom Jumlah Siswa untuk menandai hari libur.")

    df_harian_input = st.session_state.data_harian.copy()
    
    edited_harian = st.data_editor(
        df_harian_input,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        key="editor_harian_v3"
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
                    "S": "-", "I": "-", "A": "-", "Jumlah": "-",
                    "Hadir %": "-", "Tidak hadir %": "-"
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
        st.success("✅ Perhitungan Persentase Kehadiran Harian Berhasil Diperbarui!")
        st.rerun()

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

    # EXPORT EXCEL 100% PERSIS SESUAI GAMBAR DENGAN TTD & BARIS MERAH/KUNING
    def generate_excel_harian_persis_gambar(df_data, total_s, total_i, total_a, total_th, pct_h, pct_th, tempat_c, tgl_c, wali, nip, kelas_c):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Presentase Kehadiran"

        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        align_center = Alignment(horizontal="center", vertical="center")
        font_bold = Font(name="Calibri", size=11, bold=True)
        font_normal = Font(name="Calibri", size=11)

        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # 1. HEADER MULTI-LEVEL DENGAN BORDER KERAS
        ws.merge_cells("A1:A2"); ws["A1"] = "TGL"
        ws.merge_cells("B1:B2"); ws["B1"] = "Jumlah Siswa"
        
        ws.merge_cells("C1:E1"); ws["C1"] = "Tidak hadir Karena"
        ws["C2"] = "S"; ws["D2"] = "I"; ws["E2"] = "A"

        ws.merge_cells("F1:F2"); ws["F1"] = "Jumlah"

        ws.merge_cells("G1:H1"); ws["G1"] = "Presentase"
        ws["G2"] = "Hadir %"; ws["H2"] = "Tidak hadir %"

        for r in range(1, 3):
            for c in range(1, 9):
                cell = ws.cell(row=r, column=c)
                cell.alignment = align_center
                cell.font = font_bold
                cell.border = thin_border

        # 2. DATA PER HARI (1-31)
        curr_row = 3
        for idx, row in df_data.iterrows():
            tgl_val = str(row.get("TGL", idx + 1))
            js_val = str(row.get("Jumlah Siswa", "")).strip().upper()

            if js_val == "MINGGU" or str(row.get("S", "")).strip() == "-":
                ws.cell(row=curr_row, column=1, value=int(tgl_val) if tgl_val.isdigit() else tgl_val).alignment = align_center
                ws.cell(row=curr_row, column=1).font = font_bold
                ws.cell(row=curr_row, column=1).border = thin_border

                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=8)
                c_minggu = ws.cell(row=curr_row, column=2, value="MINGGU")
                c_minggu.alignment = align_center
                c_minggu.font = font_bold
                
                for col_idx in range(1, 9):
                    cell = ws.cell(row=curr_row, column=col_idx)
                    cell.fill = fill_red
                    cell.border = thin_border
            else:
                ws.cell(row=curr_row, column=1, value=int(tgl_val) if tgl_val.isdigit() else tgl_val).alignment = align_center
                ws.cell(row=curr_row, column=2, value=int(float(js_val)) if js_val.isdigit() else js_val).alignment = align_center
                ws.cell(row=curr_row, column=3, value=int(float(row.get("S", 0)))).alignment = align_center
                ws.cell(row=curr_row, column=4, value=int(float(row.get("I", 0)))).alignment = align_center
                ws.cell(row=curr_row, column=5, value=int(float(row.get("A", 0)))).alignment = align_center
                ws.cell(row=curr_row, column=6, value=int(float(row.get("Jumlah", 0)))).alignment = align_center
                ws.cell(row=curr_row, column=7, value=str(row.get("Hadir %", "100%"))).alignment = align_center
                ws.cell(row=curr_row, column=8, value=str(row.get("Tidak hadir %", "0%"))).alignment = align_center

                for col_idx in range(1, 9):
                    c_cell = ws.cell(row=curr_row, column=col_idx)
                    c_cell.font = font_bold if col_idx in [1, 2, 6, 7, 8] else font_normal
                    c_cell.border = thin_border

            curr_row += 1

        # 3. BARIS JUMLAH / TOTAL (KUNING PERSIS GAMBAR)
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=2)
        c_tot_label = ws.cell(row=curr_row, column=1, value="JUMLAH")
        c_tot_label.alignment = align_center
        
        ws.cell(row=curr_row, column=3, value=total_s).alignment = align_center
        ws.cell(row=curr_row, column=4, value=total_i).alignment = align_center
        ws.cell(row=curr_row, column=5, value=total_a).alignment = align_center
        ws.cell(row=curr_row, column=6, value=total_th).alignment = align_center
        
        c_pct_h = ws.cell(row=curr_row, column=7, value=f"{pct_h}%")
        c_pct_th = ws.cell(row=curr_row, column=8, value=f"{pct_th}%")
        
        c_pct_h.alignment = align_center
        c_pct_th.alignment = align_center
        c_pct_h.fill = fill_yellow
        c_pct_th.fill = fill_yellow

        for col_idx in range(1, 9):
            cell = ws.cell(row=curr_row, column=col_idx)
            cell.font = font_bold
            cell.border = thin_border

        # 4. TANDA TANGAN DI BAGIAN BAWAH KANAN (PERSIS GAMBAR)
        curr_row += 2
        ws.cell(row=curr_row, column=6, value=f"{tempat_c}, {tgl_c}").font = font_bold
        curr_row += 1
        ws.cell(row=curr_row, column=6, value=f"Wali Kelas  {kelas_c}").font = font_bold
        
        curr_row += 4
        c_wali_name = ws.cell(row=curr_row, column=6, value=wali)
        c_wali_name.font = Font(name="Calibri", size=11, bold=True, underline="single")
        
        curr_row += 1
        ws.cell(row=curr_row, column=6, value=f"NIP.{nip}").font = font_bold

        # Lebar Kolom Dityesuaikan
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 8
        ws.column_dimensions['E'].width = 8
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 16

        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    excel_harian_bytes = generate_excel_harian_persis_gambar(
        df_current, tot_s, tot_i, tot_a, tot_th, total_pct_h, total_pct_th,
        tempat_cetak, tgl_cetak, wali_kelas, nip_wali, kelas
    )

    st.download_button(
        label="📥 UNDUH FORMAT PERSENTASE KEHADIRAN HARIAN (EXACT SESUAI GAMBAR)",
        data=excel_harian_bytes,
        file_name=f"Persentase_Kehadiran_Harian_{kelas}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
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
