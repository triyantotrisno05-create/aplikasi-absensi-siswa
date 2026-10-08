import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Sistem Absensi & Mutasi Siswa", layout="wide")

st.title("📋 Sistem Laporan Absensi & Mutasi Siswa")
st.caption("Aplikasi Rekapitulasi Absensi Bulanan, Harian, dan Mutasi Siswa Otomatis")

# Sidebar - Informasi Sekolah & Kelas
st.sidebar.header("🏫 Data Sekolah & Kelas")
nama_sekolah = st.sidebar.text_input("Nama Sekolah", "SDN / SMPN / SMAN Negeri")
kelas = st.sidebar.text_input("Kelas", "Kelas 5A")
bulan_tahun = st.sidebar.text_input("Bulan / Periode", "Oktober 2026")
wali_kelas = st.sidebar.text_input("Nama Wali Kelas", "Guru Pembimbing, S.Pd.")

# Tab Fitur
tab1, tab2, tab3 = st.tabs(["📊 1. Rekapitulasi Absensi Bulanan", "📅 2. Kehadiran Harian Siswa", "🔄 3. Mutasi Siswa"])

# ==========================================
# TAB 1: REKAPITULASI ABSENSI BULANAN
# ==========================================
with tab1:
    st.subheader("📊 Rekapitulasi Absensi Bulanan")
    
    # --- FITUR UPLOAD FILE ---
    st.markdown("### 📤 Upload Data Siswa dari Excel / CSV")
    uploaded_file = st.file_uploader(
        "Pilih file Excel (.xlsx) atau CSV (.csv) berisi daftar siswa", 
        type=["xlsx", "csv"],
        key="uploader_bulanan"
    )

    # Inisialisasi Data Default jika belum ada
    if "data_bulanan" not in st.session_state:
        st.session_state.data_bulanan = pd.DataFrame({
            "No": [1, 2, 3],
            "NIS/NISN": ["1001", "1002", "1003"],
            "Nama Siswa": ["Ahmad Fauzi", "Budi Santoso", "Citra Dewi"],
            "Sakit (S)": [0, 0, 0],
            "Izin (I)": [0, 0, 0],
            "Alpa (A)": [0, 0, 0]
        })

    # Logika jika user mengunggah file
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df_upload = pd.read_csv(uploaded_file)
            else:
                df_upload = pd.read_excel(uploaded_file)
            
            # Pastikan kolom wajib ada
            cols_required = ["No", "NIS/NISN", "Nama Siswa"]
            for col in ["Sakit (S)", "Izin (I)", "Alpa (A)"]:
                if col not in df_upload.columns:
                    df_upload[col] = 0
            
            st.session_state.data_bulanan = df_upload
            st.success("✅ Data siswa berhasil diunggah!")
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    st.write("Atur jumlah Sakit (S), Izin (I), dan Alpa (A) pada tabel di bawah ini:")

    # Form/Editor Input Data
    edited_df = st.data_editor(
        st.session_state.data_bulanan,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_bulanan"
    )

    # Kalkulasi Otomatis
    df_calc = edited_df.copy()
    for col in ["Sakit (S)", "Izin (I)", "Alpa (A)"]:
        df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)

    df_calc["Total Ketidakhadiran"] = df_calc["Sakit (S)"] + df_calc["Izin (I)"] + df_calc["Alpa (A)"]
    
    # Asumsi Hari Efektif dalam 1 Bulan
    hari_efektif = st.number_input("Jumlah Hari Efektif Belajar (Bulan Ini)", min_value=1, value=24)
    df_calc["% Kehadiran"] = ((hari_efektif - df_calc["Total Ketidakhadiran"]) / hari_efektif * 100).round(1)
    df_calc["% Kehadiran"] = df_calc["% Kehadiran"].apply(lambda x: f"{max(0, x)}%")

    st.markdown("### 📈 Tabel Hasil Rekapitulasi")
    st.dataframe(df_calc, use_container_width=True)

    # Export Excel Terformat
    def generate_excel_rekap():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Absensi Bulanan"

        ws.append([nama_sekolah.upper()])
        ws.append([f"LAPORAN REKAPITULASI ABSENSI SISWA - {kelas.upper()}"])
        ws.append([f"PERIODE: {bulan_tahun.upper()}"])
        ws.append([])

        headers = ["No", "NIS/NISN", "Nama Siswa", "Sakit (S)", "Izin (I)", "Alpa (A)", "Total Absen", "% Kehadiran"]
        ws.append(headers)

        for _, row in df_calc.iterrows():
            ws.append([
                row.get("No", ""),
                row.get("NIS/NISN", ""),
                row.get("Nama Siswa", ""),
                row.get("Sakit (S)", 0),
                row.get("Izin (I)", 0),
                row.get("Alpa (A)", 0),
                row.get("Total Ketidakhadiran", 0),
                row.get("% Kehadiran", "100%")
            ])

        last_row = ws.max_row
        start_data_row = 5
        ws.append([
            "JUMLAH TOTAL", "", "",
            f"=SUM(D{start_data_row}:D{last_row})",
            f"=SUM(E{start_data_row}:E{last_row})",
            f"=SUM(F{start_data_row}:F{last_row})",
            f"=SUM(G{start_data_row}:G{last_row})",
            ""
        ])

        # Styling
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        title_font = Font(name="Arial", size=12, bold=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for cell in ws[1]: cell.font = title_font
        for cell in ws[2]: cell.font = title_font

        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=5, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for r in range(5, ws.max_row + 1):
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border
                if c in [1, 2, 4, 5, 6, 7, 8]:
                    cell.alignment = Alignment(horizontal="center")

        total_row = ws.max_row
        ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=total_row, column=c)
            cell.font = Font(name="Arial", size=11, bold=True)
            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        ws.append([])
        ws.append([])
        ws.append(["", "", "", "", "", "", "Wali Kelas,"])
        ws.append([])
        ws.append([])
        ws.append(["", "", "", "", "", "", f"({wali_kelas})"])

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(output)
        return output.getvalue()

    st.download_button(
        label="📥 Download Excel Rekap Bulanan",
        data=generate_excel_rekap(),
        file_name=f"Rekap_Absensi_{kelas}_{bulan_tahun}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# TAB 2: KEHADIRAN SISWA PER HARI
# ==========================================
with tab2:
    st.subheader("📅 Kehadiran Harian Siswa")
    tgl_pilih = st.date_input("Pilih Tanggal Absensi")

    if "data_harian" not in st.session_state:
        st.session_state.data_harian = pd.DataFrame({
            "No": [1, 2, 3],
            "Nama Siswa": ["Ahmad Fauzi", "Budi Santoso", "Citra Dewi"],
            "Status Kehadiran": ["Hadir", "Hadir", "Izin"],
            "Keterangan Catatan": ["-", "-", "Acara Keluarga"]
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

    summary_harian = edited_harian["Status Kehadiran"].value_counts().reset_index()
    summary_harian.columns = ["Status", "Jumlah Siswa"]
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown(f"**Ringkasan Absensi Tanggal: {tgl_pilih.strftime('%d-%m-%Y')}**")
        st.dataframe(summary_harian, use_container_width=True)
    with col_b:
        total_siswa = len(edited_harian)
        hadir_count = (edited_harian["Status Kehadiran"] == "Hadir").sum()
        persen_hadir = (hadir_count / total_siswa * 100) if total_siswa > 0 else 0
        st.metric("Persentase Kehadiran Harian", f"{persen_hadir:.1f}%")

# ==========================================
# TAB 3: MUTASI SISWA
# ==========================================
with tab3:
    st.subheader("🔄 Catatan Mutasi Siswa (Masuk / Keluar)")

    if "data_mutasi" not in st.session_state:
        st.session_state.data_mutasi = pd.DataFrame({
            "No": [1],
            "Tanggal": ["2026-10-01"],
            "Nama Siswa": ["Eko Prasetyo"],
            "Jenis Mutasi": ["Masuk"],
            "Asal / Tujuan Sekolah": ["SDN 02 Pagi"],
            "Alasan": ["Pindah Tugas Orang Tua"]
        })

    edited_mutasi = st.data_editor(
        st.session_state.data_mutasi,
        column_config={
            "Jenis Mutasi": st.column_config.SelectboxColumn(
                "Jenis Mutasi",
                options=["Masuk", "Keluar"],
                required=True
            )
        },
        num_rows="dynamic",
        use_container_width=True,
        key="editor_mutasi"
    )

    st.markdown("### 📊 Total Mutasi")
    m_masuk = (edited_mutasi["Jenis Mutasi"] == "Masuk").sum()
    m_keluar = (edited_mutasi["Jenis Mutasi"] == "Keluar").sum()
    
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("Siswa Masuk", f"{m_masuk} Orang")
    col_m2.metric("Siswa Keluar", f"{m_keluar} Orang")
