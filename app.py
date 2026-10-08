import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime, date
from PIL import Image

# ==========================================
# 1. KONFIGURASI HALAMAN & STATE
# ==========================================
st.set_page_config(page_title="Rekapitulasi Absensi Siswa", layout="wide")

if "data_absensi" not in st.session_state:
    st.session_state.data_absensi = pd.DataFrame([
        {"No": 1, "Nama Murid": "ABANG MUSHAWIR EDO", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
        {"No": 2, "Nama Murid": "AHMAD YANI", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 2},
        {"No": 3, "Nama Murid": "AL JAMI", "L/P": "L", "Nomor Induk": "", "HBE": 25, "S": 1, "I": 0, "A": 3},
        {"No": 4, "Nama Murid": "AYU NINGSIH", "L/P": "P", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 0},
        {"No": 5, "Nama Murid": "EPRI SASKIA", "L/P": "P", "Nomor Induk": "", "HBE": 25, "S": 0, "I": 0, "A": 2},
    ])

# ==========================================
# 2. SIDEBAR - DATA SEKOLAH & KELAS
# ==========================================
st.sidebar.header("🏫 Data Sekolah & Kelas")

uploaded_logo = st.sidebar.file_uploader("Upload Logo Sekolah (PNG / JPG)", type=["png", "jpg", "jpeg"])
if uploaded_logo is not None:
    logo_img = Image.open(uploaded_logo)
    st.sidebar.image(logo_img, width=120, caption="Logo Sekolah")

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

df_to_edit = st.session_state.data_absensi.copy().reset_index(drop=True)

edited_df = st.data_editor(
    df_to_edit,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key="tabel_absensi_v2"
)

st.session_state.data_absensi = edited_df.reset_index(drop=True).copy()

# ==========================================
# 5. KALKULASI LAPORAN (TANPA JUMLAH HADIR)
# ==========================================
df_calc = edited_df.copy()
for col in ["HBE", "S", "I", "A"]:
    df_calc[col] = pd.to_numeric(df_calc[col], errors="coerce").fillna(0).astype(int)

df_calc["JUMLAH"] = df_calc["S"] + df_calc["I"] + df_calc["A"]

# Kalkulasi Presentase
df_calc["PRESENTASE S"] = (df_calc["S"] / df_calc["HBE"]).fillna(0)
df_calc["PRESENTASE I"] = (df_calc["I"] / df_calc["HBE"]).fillna(0)
df_calc["PRESENTASE A"] = (df_calc["A"] / df_calc["HBE"]).fillna(0)
df_calc["PRESENTASE KEHADIRAN %"] = ((df_calc["HBE"] - df_calc["JUMLAH"]) / df_calc["HBE"]).fillna(0)

# Tampilan Tabel Web (Kolom 'JUMLAH HADIR' Dihapus)
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
# 6. EXCEL GENERATOR (TANPA JUMLAH HADIR)
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

    # Header Dokumen
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

    # Header Tabel Excel
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

    # Isi Data Tabel Excel
    start_row = 9
    num_students = len(df_calc)
    for idx, row in df_calc.iterrows():
        r = start_row + idx
        ws.cell(row=r, column=1, value=row.get("No", idx + 1))
        ws.cell(row=r, column=2, value=row.get("Nama Murid", ""))
        
        lp_cell = ws.cell(row=r, column=3, value=row.get("L/P", "L"))
        lp_cell.fill = green_fill
        
        ws.cell(row=r, column=4, value=row.get("Nomor Induk", ""))
        ws.cell(row=r, column=5, value=row.get("HBE", 25))
        ws.cell(row=r, column=6, value=row.get("S", 0))
        ws.cell(row=r, column=7, value=row.get("I", 0))
        ws.cell(row=r, column=8, value=row.get("A", 0))
        
        # Rumus Excel Jumlah Absen & Presentase
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

    # Baris Total Excel
    total_row_idx = start_row + num_students
    ws.merge_cells(f"A{total_row_idx}:B{total_row_idx}")
    cell_tot_lbl = ws.cell(row=total_row_idx, column=1, value="JUMLAH")
    cell_tot_lbl.font = font_title
    cell_tot_lbl.alignment = Alignment(horizontal="center", vertical="center")

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

    # Ringkasan L/P di Excel
    r_sum1 = total_row_idx + 2
    r_sum2 = r_sum1 + 1
    r_sum3 = r_sum2 + 1

    ws.cell(row=r_sum1, column=1, value="Laki – Laki").font = font_bold
    ws.cell(row=r_sum1, column=3, value=":").alignment = Alignment(horizontal="center")
    ws.cell(row=r_sum1, column=4, value=f'=COUNTIF(C9:C{total_row_idx-1}, "L")').font = font_bold

    ws.cell(row=r_sum2, column=1, value="Perempuan").font = font_bold
    ws.cell(row=r_sum2, column=3, value=":").alignment = Alignment(horizontal="center")
    ws.cell(row=r_sum2, column=4, value=f'=COUNTIF(C9:C{total_row_idx-1}, "P")').font = font_bold

    ws.cell(row=r_sum3, column=1, value="Jumlah akhir bulan").font = font_bold
    ws.cell(row=r_sum3, column=3, value=":").alignment = Alignment(horizontal="center")
    ws.cell(row=r_sum3, column=4, value=f'=COUNTA(C9:C{total_row_idx-1})').font = font_bold

    # Tanda Tangan Wali Kelas
    r_ttd_tgl = total_row_idx + 4
    r_ttd_jab = r_ttd_tgl + 1
    r_ttd_nama = r_ttd_jab + 4
    r_ttd_nip = r_ttd_nama + 1

    ws.cell(row=r_ttd_tgl, column=10, value=tgl_cetak_str).font = font_data
    ws.cell(row=r_ttd_jab, column=10, value=f"Wali Kelas {kelas}").font = font_data
    ws.cell(row=r_ttd_nama, column=10, value=wali_kelas).font = Font(name="Calibri", size=10, bold=True, underline="single")
    ws.cell(row=r_ttd_nip, column=10, value=f"NIP. {nip_wali}" if nip_wali else "").font = font_bold

    # Lebar Kolom Excel
    col_widths = {
        'A': 5, 'B': 30, 'C': 6, 'D': 16, 'E': 6,
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
