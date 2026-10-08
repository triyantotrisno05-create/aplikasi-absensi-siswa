import streamlit as st
import pandas as pd
import io
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(page_title="Sistem Laporan Absensi Siswa", layout="wide")

st.title("📊 Aplikasi Laporan Absensi Bulanan Siswa")

# --- SIDEBAR PENGATURAN ---
st.sidebar.header("⚙️ Pengaturan Laporan")
nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SMP NEGERI 1 NANGA MAHAP")
tahun_ajaran = st.sidebar.text_input("Tahun Pelajaran", "2025/2026")
kelas = st.sidebar.text_input("Kelas", "IX C")

# Input Tanggal Lengkap (Tanggal, Bulan, Tahun)
tgl_laporan = st.sidebar.date_input("Tanggal Laporan", datetime.date.today())
# Format Tanggal ke Bahasa Indonesia (contoh: 25 April 2026)
bulan_indo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
tgl_formatted = f"{tgl_laporan.day} {bulan_indo[tgl_laporan.month - 1]} {tgl_laporan.year}"

hbe = st.sidebar.number_input("Hari Belajar Efektif (HBE)", min_value=1, value=25)
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "Triyanto trisno, S.Pd")
nip_wali = st.sidebar.text_input("NIP Wali Kelas", "NIP.1993052020242110006")

# --- DATA INITIALIZATION ---
if 'data_absensi' not in st.session_state:
    st.session_state.data_absensi = pd.DataFrame([
        {"Nama Murid": "ABANG MUSHAWIR EDO", "L/P": "L", "Sakit (S)": 0, "Izin (I)": 0, "Alpha (A)": 0},
        {"Nama Murid": "AHMAD YANI", "L/P": "L", "Sakit (S)": 0, "Izin (I)": 0, "Alpha (A)": 2},
        {"Nama Murid": "AL JAMI", "L/P": "L", "Sakit (S)": 1, "Izin (I)": 0, "Alpha (A)": 3},
        {"Nama Murid": "AYU NINGSIH", "L/P": "P", "Sakit (S)": 0, "Izin (I)": 0, "Alpha (A)": 0},
        {"Nama Murid": "EPRI SASKIA", "L/P": "P", "Sakit (S)": 0, "Izin (I)": 0, "Alpha (A)": 2},
    ])

if 'data_mutasi' not in st.session_state:
    st.session_state.data_mutasi = pd.DataFrame([
        {"Nama Siswa": "", "NIS/NISN": "", "L/P": "L", "Agama": "", "Umur": "", "Pekerjaan Orang Tua": "", "Tgl. Keluar": "", "Tgl. Masuk": ""}
    ])

tab1, tab2, tab3 = st.tabs(["📝 Data & Absensi Siswa", "🔄 Mutasi Siswa", "📥 Unduh Laporan Excel"])

def safe_int(val):
    try:
        if pd.isna(val) or val is None or str(val).strip() == "":
            return 0
        return int(float(val))
    except:
        return 0

# --- TAB 1: DATA SISWA & ABSENSI ---
with tab1:
    st.subheader("📋 Input & Edit Data Absensi Siswa")
    st.info(f"📅 **Tanggal Laporan Dipilih:** {tgl_formatted}")
    
    with st.form("form_absensi"):
        df_edited = st.data_editor(
            st.session_state.data_absensi,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "Nama Murid": st.column_config.TextColumn("Nama Murid", required=True),
                "L/P": st.column_config.SelectboxColumn("L/P", options=["L", "P"], required=True, default="L"),
                "Sakit (S)": st.column_config.NumberColumn("Sakit (S)", min_value=0, default=0),
                "Izin (I)": st.column_config.NumberColumn("Izin (I)", min_value=0, default=0),
                "Alpha (A)": st.column_config.NumberColumn("Alpha (A)", min_value=0, default=0),
            },
            key="editor_absensi"
        )
        submit_btn = st.form_submit_button("💾 Simpan & Perbarui Rekap")
        if submit_btn:
            st.session_state.data_absensi = df_edited
            st.success("Data berhasil diperbarui!")

    df_curr = st.session_state.data_absensi.copy()
    
    df_curr['Sakit (S)'] = df_curr['Sakit (S)'].apply(safe_int)
    df_curr['Izin (I)'] = df_curr['Izin (I)'].apply(safe_int)
    df_curr['Alpha (A)'] = df_curr['Alpha (A)'].apply(safe_int)
    df_curr['Total Absen'] = df_curr['Sakit (S)'] + df_curr['Izin (I)'] + df_curr['Alpha (A)']
    df_curr['Total Hadir'] = df_curr.apply(lambda r: max(0, hbe - r['Total Absen']), axis=1)

    tot_s = df_curr['Sakit (S)'].sum()
    tot_i = df_curr['Izin (I)'].sum()
    tot_a = df_curr['Alpha (A)'].sum()
    tot_absen_all = tot_s + tot_i + tot_a
    tot_hadir_all = df_curr['Total Hadir'].sum()
    tot_kesempatan_hadir = len(df_curr) * hbe
    persen_hadir_total = (tot_hadir_all / tot_kesempatan_hadir * 100) if tot_kesempatan_hadir > 0 else 0

    st.markdown("---")
    st.subheader("📊 Rekapitulasi Otomatis")
    summary_df = pd.DataFrame([{
        "Total Siswa": len(df_curr),
        "Laki-laki (L)": len(df_curr[df_curr['L/P'] == 'L']),
        "Perempuan (P)": len(df_curr[df_curr['L/P'] == 'P']),
        "Total Sakit (S)": tot_s,
        "Total Izin (I)": tot_i,
        "Total Alpha (A)": tot_a,
        "Total Absen": tot_absen_all,
        "Total Hadir": tot_hadir_all,
        "Persentase Kehadiran Kelas": f"{persen_hadir_total:.2f}%"
    }])
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

# --- TAB 2: MUTASI SISWA ---
with tab2:
    st.subheader("🔄 Data Mutasi Siswa (Masuk / Keluar)")
    with st.form("form_mutasi"):
        df_mutasi_edited = st.data_editor(
            st.session_state.data_mutasi,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_mutasi"
        )
        submit_mutasi = st.form_submit_button("💾 Simpan Data Mutasi")
        if submit_mutasi:
            st.session_state.data_mutasi = df_mutasi_edited
            st.success("Data mutasi berhasil diperbarui!")

# --- TAB 3: UNDUH EXCEL ---
with tab3:
    st.subheader("📥 Export ke Format Excel")
    
    def generate_excel():
        output = io.BytesIO()
        wb = Workbook()
        
        font_title = Font(name="Calibri", size=12, bold=True)
        font_header = Font(name="Calibri", size=10, bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")

        ws = wb.active
        ws.title = "REKAP ABSEN"
        
        ws["A1"] = "REKAPITULASI ABSENSI SISWA"
        ws["A1"].font = font_title
        ws["A2"] = nama_sekolah
        ws["A2"].font = font_title
        ws["A3"] = f"TAHUN PELAJARAN {tahun_ajaran}"
        ws["A3"].font = font_title
        ws["A4"] = f"KELAS : {kelas}"
        ws["A4"].font = font_header

        headers = ["NO", "NAMA MURID", "L/P", "HBE", "SAKIT (S)", "IZIN (I)", "ALPHA (A)", "JUMLAH ABSEN", "JUMLAH HADIR", "PRESENTASE HADIR"]
        ws.append([])
        ws.append(headers)
        header_row = 6
        
        for col in range(1, 11):
            cell = ws.cell(row=header_row, column=col)
            cell.font = font_header
            cell.alignment = align_center
            cell.border = thin_border

        df = st.session_state.data_absensi
        start_row = header_row + 1
        
        for idx, r in df.iterrows():
            row_num = start_row + idx
            s = safe_int(r.get('Sakit (S)'))
            i = safe_int(r.get('Izin (I)'))
            a = safe_int(r.get('Alpha (A)'))
            
            nama = str(r.get('Nama Murid', '')) if pd.notnull(r.get('Nama Murid')) else ""
            lp = str(r.get('L/P', 'L')) if pd.notnull(r.get('L/P')) else "L"
            
            ws.cell(row=row_num, column=1, value=idx+1)
            ws.cell(row=row_num, column=2, value=nama)
            ws.cell(row=row_num, column=3, value=lp)
            ws.cell(row=row_num, column=4, value=hbe)
            ws.cell(row=row_num, column=5, value=s)
            ws.cell(row=row_num, column=6, value=i)
            ws.cell(row=row_num, column=7, value=a)
            ws.cell(row=row_num, column=8, value=f"=SUM(E{row_num}:G{row_num})")
            ws.cell(row=row_num, column=9, value=f"=MAX(0, D{row_num}-H{row_num})")
            
            cell_p = ws.cell(row=row_num, column=10, value=f"=I{row_num}/D{row_num}")
            cell_p.number_format = "0.0%"
            
            for col in range(1, 11):
                c = ws.cell(row=row_num, column=col)
                c.border = thin_border
                c.alignment = align_center if col != 2 else align_left

        end_row = start_row + len(df) - 1
        total_row = end_row + 1

        # Baris Rekapitulasi Total di Bawah Tabel
        ws.cell(row=total_row, column=2, value="JUMLAH TOTAL").font = font_header
        ws.cell(row=total_row, column=4, value=f"=SUM(D{start_row}:D{end_row})").font = font_header
        ws.cell(row=total_row, column=5, value=f"=SUM(E{start_row}:E{end_row})").font = font_header
        ws.cell(row=total_row, column=6, value=f"=SUM(F{start_row}:F{end_row})").font = font_header
        ws.cell(row=total_row, column=7, value=f"=SUM(G{start_row}:G{end_row})").font = font_header
        ws.cell(row=total_row, column=8, value=f"=SUM(H{start_row}:H{end_row})").font = font_header
        ws.cell(row=total_row, column=9, value=f"=SUM(I{start_row}:I{end_row})").font = font_header
        
        cell_tot_p = ws.cell(row=total_row, column=10, value=f"=I{total_row}/D{total_row}")
        cell_tot_p.font = font_header
        cell_tot_p.number_format = "0.0%"

        for col in range(1, 11):
            c = ws.cell(row=total_row, column=col)
            c.border = thin_border
            c.alignment = align_center if col != 2 else align_left

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        # Tanda Tangan dengan Tanggal Lengkap
        ttd_row = total_row + 3
        lokasi = nama_sekolah.split()[-1] if len(nama_sekolah.split()) > 0 else "Nanga Mahap"
        ws.cell(row=ttd_row, column=8, value=f"{lokasi}, {tgl_formatted}")
        ws.cell(row=ttd_row+1, column=8, value=f"Wali Kelas {kelas}")
        ws.cell(row=ttd_row+4, column=8, value=wali_kelas).font = font_header
        ws.cell(row=ttd_row+5, column=8, value=nip_wali)

        wb.save(output)
        output.seek(0)
        return output

    excel_file = generate_excel()
    st.download_button(
        label="📥 Download File Excel Rekap Bulanan",
        data=excel_file,
        file_name=f"ABSENSI_{kelas}_{tgl_formatted.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )