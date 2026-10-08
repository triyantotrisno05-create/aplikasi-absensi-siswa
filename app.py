import calendar
from datetime import datetime
import pandas as pd
import streamlit as st

# ==========================================
# 1. KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Sistem Rekapitulasi Absensi Sekolah",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS Styling
st.markdown(
    """
    <style>
    .title-header {
        text-align: center;
        font-weight: 800;
        font-size: 2.0rem;
        color: #212529;
        margin-bottom: 0px;
    }
    .subtitle-header {
        text-align: center;
        font-weight: 700;
        font-size: 1.1rem;
        color: #495057;
        margin-bottom: 15px;
    }
    div.stButton > button:first-child, div.stDownloadButton > button:first-child {
        background-color: #ff4d4f !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-size: 14px !important;
        width: 100%;
        margin-top: 5px;
    }
    div.stButton > button:first-child:hover, div.stDownloadButton > button:first-child:hover {
        background-color: #ff2a2d !important;
        color: white !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. PENGATURAN IDENTITAS SEKOLAH (SIDEBAR)
# ==========================================
with st.sidebar:
    st.header("⚙️ Pengaturan Header & Data")

    nama_sekolah = st.text_input(
        "Nama Sekolah", value="SMP NEGERI 1 NANGA MAHAP"
    )
    kelas_input = st.text_input("Kelas", value="IX C")
    tahun_ajaran = st.text_input("Tahun Pelajaran", value="2025/2026")
    nama_wali = st.text_input("Nama Wali Kelas", value="Trivanto trisno, S.Pd")
    nip_wali = st.text_input("NIP Wali Kelas", value="199305202024211000")
    kota_lokasi = st.text_input("Kota / Kecamatan", value="Nanga Mahap")

    st.subheader("🖼️ Upload Logo Sekolah")
    uploaded_logo = st.file_uploader(
        "Upload Logo (PNG / JPG)", type=["png", "jpg", "jpeg"]
    )

# ==========================================
# 3. TAMPILAN HEADER UTAMA APLIKASI
# ==========================================
col_logo1, col_text, col_logo2 = st.columns([1, 4, 1])

with col_logo1:
    if uploaded_logo is not None:
        st.image(uploaded_logo, width=90)
    else:
        # Placeholder jika belum upload logo
        st.image(
            "https://cdn-icons-png.flaticon.com/512/2991/2991148.png", width=90
        )

with col_text:
    st.markdown(
        f"<h1 class='title-header'>{nama_sekolah.upper()}</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<h3 class='subtitle-header'>REKAPITULASI ABSENSI & MUTASI SISWA - {kelas_input.upper()} ({tahun_ajaran})</h3>",
        unsafe_allow_html=True,
    )

with col_logo2:
    pass

st.divider()

# ==========================================
# 4. TAB NAVIGASI UTAMA
# ==========================================
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "📊 Rekap Absensi Bulanan",
        "📅 Persentase Kehadiran Per Hari",
        "📂 Upload Data Siswa",
        "🔄 Mutasi Siswa",
    ]
)

nama_bulan = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
]

# ------------------------------------------
# TAB 1: REKAP ABSENSI BULANAN
# ------------------------------------------
with tab1:
    st.subheader("📊 Rekap Absensi Bulanan")
    st.caption("Lakukan pengeditan data absensi bulanan siswa.")

    col_b1, col_b2, col_b3 = st.columns([2, 2, 2])
    with col_b1:
        bln_rekap = st.selectbox(
            "Pilih Bulan Rekap",
            options=list(range(1, 13)),
            format_func=lambda x: nama_bulan[x - 1],
            key="bln_tab1",
        )
    with col_b2:
        thn_rekap = st.number_input(
            "Pilih Tahun Rekap",
            value=datetime.now().year,
            key="thn_tab1",
        )
    with col_b3:
        hbe_input = st.number_input(
            "Hari Belajar Efektif (HBE)", min_value=1, max_value=31, value=25
        )

    if "df_siswa" in st.session_state and not st.session_state[
        "df_siswa"
    ].empty:
        df_base = st.session_state["df_siswa"].copy()
    else:
        default_data = [
            ("1", "ABANG MUSHAVIR EDO", "L", "148355770"),
            ("2", "AHMAD YANI", "L", "3143662047"),
            ("3", "AL JAMI", "L", "3137752600"),
            ("4", "AYU NINGSIH", "P", "3139019097"),
            ("5", "EPRI SASKIA", "P", "3131994928"),
        ]
        df_base = pd.DataFrame(
            default_data, columns=["NO", "NAMA MURID", "L/P", "NOMOR INDUK"]
        )

    session_key_rekap = f"df_rekap_{bln_rekap}_{thn_rekap}"

    if (
        session_key_rekap not in st.session_state
        or st.session_state.get("need_reload_rekap", False)
    ):
        df_init = df_base.copy()
        df_init["HBE"] = hbe_input
        if "S" not in df_init.columns:
            df_init["S"] = 0
        if "I" not in df_init.columns:
            df_init["I"] = 0
        if "A" not in df_init.columns:
            df_init["A"] = 0
        st.session_state[session_key_rekap] = df_init
        st.session_state["need_reload_rekap"] = False

    def hitung_rekap_bulanan(df, hbe_val):
        df_calc = df.copy()
        jml_absen_list, jml_hadir_list = [], []
        pct_s_list, pct_i_list, pct_a_list, pct_hadir_list = [], [], [], []

        for _, row in df_calc.iterrows():
            s = int(row["S"]) if str(row.get("S", 0)).isdigit() else 0
            i = int(row["I"]) if str(row.get("I", 0)).isdigit() else 0
            a = int(row["A"]) if str(row.get("A", 0)).isdigit() else 0

            tot_absen = s + i + a
            tot_hadir = max(0, hbe_val - tot_absen)

            pct_s = round((s / hbe_val) * 100) if hbe_val > 0 else 0
            pct_i = round((i / hbe_val) * 100) if hbe_val > 0 else 0
            pct_a = round((a / hbe_val) * 100) if hbe_val > 0 else 0
            pct_hadir = (
                round((tot_hadir / hbe_val) * 100) if hbe_val > 0 else 0
            )

            jml_absen_list.append(tot_absen)
            jml_hadir_list.append(tot_hadir)
            pct_s_list.append(f"{pct_s}%")
            pct_i_list.append(f"{pct_i}%")
            pct_a_list.append(f"{pct_a}%")
            pct_hadir_list.append(f"{pct_hadir}%")

        df_calc["HBE"] = hbe_val
        df_calc["JUMLAH ABSEN"] = jml_absen_list
        df_calc["JUMLAH HADIR"] = jml_hadir_list
        df_calc["PRESENTASE S"] = pct_s_list
        df_calc["PRESENTASE I"] = pct_i_list
        df_calc["PRESENTASE A"] = pct_a_list
        df_calc["PRESENTASE KEHADIRAN"] = pct_hadir_list

        return df_calc

    df_rekap_edited = st.data_editor(
        st.session_state[session_key_rekap],
        key=f"editor_{session_key_rekap}",
        use_container_width=True,
        column_config={
            "NO": st.column_config.TextColumn("NO"),
            "NAMA MURID": st.column_config.TextColumn("NAMA MURID"),
            "L/P": st.column_config.TextColumn("L/P"),
            "NOMOR INDUK": st.column_config.TextColumn("NOMOR INDUK"),
            "HBE": st.column_config.NumberColumn("HBE"),
            "S": st.column_config.NumberColumn("S (Sakit)", min_value=0),
            "I": st.column_config.NumberColumn("I (Izin)", min_value=0),
            "A": st.column_config.NumberColumn("A (Alpha)", min_value=0),
        },
        hide_index=True,
    )

    df_final_rekap = hitung_rekap_bulanan(df_rekap_edited, hbe_input)

    def generate_excel_html(df_data, bulan, tahun, hbe):
        tot_l = len(
            df_data[
                df_data["L/P"].astype(str).str.upper().str.startswith("L")
            ]
        )
        tot_p = len(
            df_data[
                df_data["L/P"].astype(str).str.upper().str.startswith("P")
            ]
        )
        tot_siswa = len(df_data)

        tot_s = sum([int(x) for x in df_data["S"]])
        tot_i = sum([int(x) for x in df_data["I"]])
        tot_a = sum([int(x) for x in df_data["A"]])
        tot_absen = sum([int(x) for x in df_data["JUMLAH ABSEN"]])

        html = f"""
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head><meta charset="utf-8"/></head>
        <body>
            <h2 align="center">REKAPITULASI ABSENSI BULANAN SISWA</h2>
            <h3 align="center">{nama_sekolah.upper()} - KELAS {kelas_input.upper()}</h3>
            <p><b>Bulan:</b> {nama_bulan[bulan-1]} {tahun} | <b>HBE:</b> {hbe} Hari</p>
            <table border="1" style="border-collapse:collapse; text-align:center;">
                <thead>
                    <tr style="background-color:#d9d9d9;">
                        <th rowspan="2">NO</th>
                        <th rowspan="2">NAMA MURID</th>
                        <th rowspan="2">L/P</th>
                        <th rowspan="2">NOMOR INDUK</th>
                        <th rowspan="2">HBE</th>
                        <th colspan="3">ABSENSI</th>
                        <th rowspan="2">JUMLAH</th>
                        <th rowspan="2">JUMLAH HADIR</th>
                        <th colspan="3">PRESENTASE</th>
                        <th rowspan="2">PRESENTASE KEHADIRAN</th>
                    </tr>
                    <tr style="background-color:#d9d9d9;">
                        <th>S</th><th>I</th><th>A</th>
                        <th>S</th><th>I</th><th>A</th>
                    </tr>
                </thead>
                <tbody>
        """
        for _, r in df_data.iterrows():
            html += f"""
                    <tr>
                        <td>{r['NO']}</td>
                        <td align="left">{r['NAMA MURID']}</td>
                        <td>{r['L/P']}</td>
                        <td style="mso-number-format:'\@';">{r['NOMOR INDUK']}</td>
                        <td>{r['HBE']}</td>
                        <td>{r['S']}</td>
                        <td>{r['I']}</td>
                        <td>{r['A']}</td>
                        <td>{r['JUMLAH ABSEN']}</td>
                        <td>{r['JUMLAH HADIR']}</td>
                        <td>{r['PRESENTASE S']}</td>
                        <td>{r['PRESENTASE I']}</td>
                        <td>{r['PRESENTASE A']}</td>
                        <td><b>{r['PRESENTASE KEHADIRAN']}</b></td>
                    </tr>
            """

        html += f"""
                    <tr style="background-color:#ffff00; font-weight:bold;">
                        <td colspan="4">JUMLAH</td>
                        <td>{tot_siswa}</td>
                        <td>{tot_s}</td>
                        <td>{tot_i}</td>
                        <td>{tot_a}</td>
                        <td>{tot_absen}</td>
                        <td colspan="5"></td>
                    </tr>
                </tbody>
            </table>
            <br/>
            <table>
                <tr><td><b>Laki - Laki</b></td><td>: {tot_l}</td></tr>
                <tr><td><b>Perempuan</b></td><td>: {tot_p}</td></tr>
                <tr><td><b>Jumlah akhir bulan</b></td><td>: {tot_siswa}</td></tr>
            </table>
            <br/><br/>
            <table width="100%">
                <tr>
                    <td width="60%"></td>
                    <td align="center">
                        {kota_lokasi}, 30 {nama_bulan[bulan-1]} {tahun}<br/>
                        Wali Kelas {kelas_input}<br/><br/><br/><br/>
                        <b><u>{nama_wali}</u></b><br/>
                        NIP. {nip_wali}
                    </td>
                </tr>
            </table>
        </body>
        </html>
        """
        return html.encode("utf-8")

    excel_bytes = generate_excel_html(
        df_final_rekap, bln_rekap, thn_rekap, hbe_input
    )

    st.download_button(
        label="💾 UNDUH FORMAT REKAPITULASI BULANAN (EXACT PERSIS EXCEL)",
        data=excel_bytes,
        file_name=f"Rekapitulasi_Bulanan_{nama_bulan[bln_rekap-1]}_{thn_rekap}.xls",
        mime="application/vnd.ms-excel",
        key="btn_download_bulanan",
    )


# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN PER HARI
# ------------------------------------------
with tab2:
    st.subheader("📅 Persentase Kehadiran Per Hari (Format Persis Foto Excel)")
    st.caption("Isi nilai S, I, A atau tulis 'MINGGU' pada kolom Jumlah Siswa.")

    col1, col2, _ = st.columns([2, 2, 4])
    with col1:
        bulan_selected = st.selectbox(
            "Pilih Bulan",
            options=list(range(1, 13)),
            format_func=lambda x: f"{nama_bulan[x-1]} (Bulan {x})",
            index=3,
            key="bln_tab2",
        )
    with col2:
        tahun_selected = st.number_input(
            "Pilih Tahun",
            min_value=2020,
            max_value=2035,
            value=2026,
            key="thn_tab2",
        )

    _, total_hari = calendar.monthrange(tahun_selected, bulan_selected)
    session_key = f"df_harian_{bulan_selected}_{tahun_selected}"

    if session_key not in st.session_state:
        data_harian = []
        for tgl in range(1, total_hari + 1):
            is_minggu = (
                calendar.weekday(tahun_selected, bulan_selected, tgl) == 6
            )
            data_harian.append(
                {
                    "TGL": tgl,
                    "Jumlah Siswa": "MINGGU" if is_minggu else "30",
                    "S": 0,
                    "I": 0,
                    "A": 0,
                    "Jumlah": 0,
                    "Hadir %": "-" if is_minggu else "100%",
                    "Tidak hadir %": "-" if is_minggu else "0%",
                }
            )
        st.session_state[session_key] = pd.DataFrame(data_harian)

    def hitung_ulang(df):
        df_copy = df.copy()
        jml_list, h_list, th_list = [], [], []

        for _, row in df_copy.iterrows():
            jml_str = str(row["Jumlah Siswa"]).strip()
            s = int(row["S"]) if str(row["S"]).isdigit() else 0
            i = int(row["I"]) if str(row["I"]).isdigit() else 0
            a = int(row["A"]) if str(row["A"]).isdigit() else 0

            tot_absen = s + i + a
            jml_list.append(tot_absen)

            if (
                jml_str.upper() == "MINGGU"
                or not jml_str.isdigit()
                or int(jml_str) == 0
            ):
                h_list.append("-")
                th_list.append("-")
            else:
                tot_siswa = int(jml_str)
                hp = max(
                    0.0, round(((tot_siswa - tot_absen) / tot_siswa) * 100)
                )
                thp = min(100.0, round((tot_absen / tot_siswa) * 100))
                h_list.append(f"{int(hp)}%")
                th_list.append(f"{int(thp)}%")

        df_copy["Jumlah"] = jml_list
        df_copy["Hadir %"] = h_list
        df_copy["Tidak hadir %"] = th_list
        return df_copy

    edited_df = st.data_editor(
        st.session_state[session_key],
        key=f"editor_{session_key}",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
    )

    def generate_excel_harian_html(df_data, bulan, tahun):
        df_calc = hitung_ulang(df_data)

        tot_s = sum(
            [
                int(x)
                for x in df_calc["S"]
                if str(x).isdigit()
                and str(df_calc.loc[_].get("Jumlah Siswa")).upper() != "MINGGU"
            ]
        )
        tot_i = sum(
            [
                int(x)
                for x in df_calc["I"]
                if str(x).isdigit()
                and str(df_calc.loc[_].get("Jumlah Siswa")).upper() != "MINGGU"
            ]
        )
        tot_a = sum(
            [
                int(x)
                for x in df_calc["A"]
                if str(x).isdigit()
                and str(df_calc.loc[_].get("Jumlah Siswa")).upper() != "MINGGU"
            ]
        )
        tot_jumlah = tot_s + tot_i + tot_a

        valid_rows = df_calc[df_calc["Hadir %"] != "-"]
        if len(valid_rows) > 0:
            avg_hadir = round(
                sum(
                    [
                        float(x.replace("%", ""))
                        for x in valid_rows["Hadir %"]
                    ]
                )
                / len(valid_rows)
            )
            avg_thadir = round(
                sum(
                    [
                        float(x.replace("%", ""))
                        for x in valid_rows["Tidak hadir %"]
                    ]
                )
                / len(valid_rows)
            )
        else:
            avg_hadir, avg_thadir = 100, 0

        html = f"""
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <meta charset="utf-8"/>
            <style>
                body {{ font-family: 'Calibri', 'Segoe UI', Arial, sans-serif; }}
                table {{ border-collapse: collapse; width: 100%; font-size: 11pt; }}
                th, td {{ border: 1px solid black; padding: 4px; text-align: center; }}
                .bg-minggu {{ background-color: #ff0000; color: black; font-weight: bold; text-align: center; }}
                .bg-jumlah {{ font-weight: bold; background-color: #ffffff; }}
                .bg-kuning {{ background-color: #ffff00; font-weight: bold; }}
            </style>
        </head>
        <body>
            <table>
                <thead>
                    <tr style="font-weight:bold; background-color:#ffffff;">
                        <th rowspan="2" style="width: 50px;">TGL</th>
                        <th rowspan="2" style="width: 100px;">Jumlah Siswa</th>
                        <th colspan="3">Tidak hadir Karena</th>
                        <th rowspan="2" style="width: 80px;">Jumlah</th>
                        <th colspan="2">Presentase</th>
                    </tr>
                    <tr style="font-weight:bold; background-color:#ffffff;">
                        <th style="width: 50px;">S</th>
                        <th style="width: 50px;">I</th>
                        <th style="width: 50px;">A</th>
                        <th style="width: 80px;">Hadir %</th>
                        <th style="width: 100px;">Tidak hadir %</th>
                    </tr>
                </thead>
                <tbody>
        """

        for idx, r in df_calc.iterrows():
            jml_siswa_str = str(r["Jumlah Siswa"]).strip()
            if jml_siswa_str.upper() == "MINGGU":
                html += f"""
                    <tr>
                        <td><b>{r['TGL']}</b></td>
                        <td colspan="7" class="bg-minggu">MINGGU</td>
                    </tr>
                """
            else:
                html += f"""
                    <tr>
                        <td><b>{r['TGL']}</b></td>
                        <td>{r['Jumlah Siswa']}</td>
                        <td>{r['S']}</td>
                        <td>{r['I']}</td>
                        <td>{r['A']}</td>
                        <td>{r['Jumlah']}</td>
                        <td>{r['Hadir %']}</td>
                        <td>{r['Tidak hadir %']}</td>
                    </tr>
                """

        html += f"""
                    <tr class="bg-jumlah">
                        <td colspan="2" style="text-align:center;"><b>JUMLAH</b></td>
                        <td><b>{tot_s}</b></td>
                        <td><b>{tot_i}</b></td>
                        <td><b>{tot_a}</b></td>
                        <td><b>{tot_jumlah}</b></td>
                        <td class="bg-kuning"><b>{avg_hadir}%</b></td>
                        <td class="bg-kuning"><b>{avg_thadir}%</b></td>
                    </tr>
                </tbody>
            </table>
            <br/><br/>
            <table style="border:none; width:100%;">
                <tr style="border:none;">
                    <td style="border:none; width:50%;"></td>
                    <td style="border:none; text-align:center;">
                        {kota_lokasi}, 30 {nama_bulan[bulan-1]} {tahun}<br/>
                        Wali Kelas {kelas_input}<br/><br/><br/><br/>
                        <b><u>{nama_wali}</u></b><br/>
                        NIP. {nip_wali}
                    </td>
                </tr>
            </table>
        </body>
        </html>
        """
        return html.encode("utf-8")

    btn_col1, btn_col2 = st.columns([1, 1.5])
    with btn_col1:
        if st.button("💾 Simpan & Hitung Ulang Persentase", key="btn_save"):
            st.session_state[session_key] = hitung_ulang(edited_df)
            st.success("✅ Data berhasil dihitung ulang dan disimpan!")
            st.rerun()

    with btn_col2:
        excel_harian_bytes = generate_excel_harian_html(
            edited_df, bulan_selected, tahun_selected
        )

        st.download_button(
            label="💾 UNDUH PERSENTASE KEHADIRAN HARIAN (PERSIS FOTO EXCEL)",
            data=excel_harian_bytes,
            file_name=f"Rekap_Kehadiran_Harian_{nama_bulan[bulan_selected-1]}_{tahun_selected}.xls",
            mime="application/vnd.ms-excel",
            key="btn_download_harian_exact",
        )


# ------------------------------------------
# TAB 3: UPLOAD DATA SISWA
# ------------------------------------------
with tab3:
    st.subheader(f"📂 Upload Data Siswa Kelas {kelas_input}")
    st.caption("Unggah file Excel (.xlsx / .xls) atau CSV daftar siswa Anda.")

    uploaded_file = st.file_uploader(
        "Pilih File Excel atau CSV", type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df_uploaded = pd.read_csv(uploaded_file, sep=";", dtype=str)
                if len(df_uploaded.columns) == 1:
                    uploaded_file.seek(0)
                    df_uploaded = pd.read_csv(uploaded_file, sep=",", dtype=str)
            else:
                df_uploaded = pd.read_excel(uploaded_file, dtype=str)

            df_uploaded.columns = [
                str(col).strip().upper() for col in df_uploaded.columns
            ]

            rename_dict = {}
            for col in df_uploaded.columns:
                if any(
                    k in col
                    for k in ["NAMA", "NAMA SISWA", "MURID", "NAMA MURID"]
                ):
                    rename_dict[col] = "NAMA MURID"
                elif any(
                    k in col
                    for k in [
                        "INDUK",
                        "NOMOR INDUK",
                        "NIS",
                        "NISN",
                        "NIPD",
                        "NO INDUK",
                    ]
                ):
                    rename_dict[col] = "NOMOR INDUK"
                elif col in ["JK", "L/P", "JENIS KELAMIN", "SEX"]:
                    rename_dict[col] = "L/P"
                elif col in ["NO", "NO.", "NOMOR"]:
                    rename_dict[col] = "NO"

            df_uploaded = df_uploaded.rename(columns=rename_dict)

            if "NO" not in df_uploaded.columns:
                df_uploaded.insert(
                    0, "NO", [str(i + 1) for i in range(len(df_uploaded))]
                )
            if "L/P" not in df_uploaded.columns:
                df_uploaded["L/P"] = "L"
            if "NOMOR INDUK" not in df_uploaded.columns:
                df_uploaded["NOMOR INDUK"] = "-"

            df_uploaded = df_uploaded[
                ["NO", "NAMA MURID", "L/P", "NOMOR INDUK"]
            ]
            df_uploaded = df_uploaded.dropna(subset=["NAMA MURID"]).reset_index(
                drop=True
            )

            st.session_state["df_siswa"] = df_uploaded
            st.session_state["need_reload_rekap"] = True
            st.success(
                f"✅ BERHASIL! Data {len(df_uploaded)} siswa telah dimasukkan!"
            )

        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    if "df_siswa" in st.session_state and not st.session_state[
        "df_siswa"
    ].empty:
        st.write("---")
        st.write("### Data Siswa Tersimpan Saat Ini:")
        st.dataframe(
            st.session_state["df_siswa"],
            use_container_width=True,
            hide_index=True,
        )


# ------------------------------------------
# TAB 4: MUTASI SISWA
# ------------------------------------------
with tab4:
    st.subheader("🔄 Data Mutasi Siswa")
    st.caption("Isi/edit tabel mutasi di bawah ini.")

    col_m_b, col_m_t = st.columns(2)
    with col_m_b:
        bln_mutasi = st.selectbox(
            "Pilih Bulan Mutasi",
            options=list(range(1, 13)),
            format_func=lambda x: nama_bulan[x - 1],
            index=3,
            key="bln_tab4",
        )
    with col_m_t:
        thn_mutasi = st.number_input(
            "Pilih Tahun Mutasi",
            value=2026,
            key="thn_tab4",
        )

    if "df_mutasi_exact" not in st.session_state:
        default_mutasi_data = []
        for i in range(1, 11):
            default_mutasi_data.append(
                {
                    "No": str(i),
                    "Nama Siswa": "" if i > 1 else "RIAN HIDAYAT",
                    "NIS / NISN": "" if i > 1 else "0148355770",
                    "L/P": "" if i > 1 else "L",
                    "Agama": "" if i > 1 else "Islam",
                    "Umur": "" if i > 1 else "15",
                    "Pekerjaan Orang Tua": "" if i > 1 else "Petani",
                    "Tgl. Keluar": "",
                    "Tgl. Masuk": "" if i > 1 else "2026-04-10",
                }
            )
        st.session_state["df_mutasi_exact"] = pd.DataFrame(default_mutasi_data)

    df_mutasi_edited = st.data_editor(
        st.session_state["df_mutasi_exact"],
        key="editor_mutasi_exact",
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "No": st.column_config.TextColumn("No"),
            "Nama Siswa": st.column_config.TextColumn("Nama Siswa"),
            "NIS / NISN": st.column_config.TextColumn("NIS / NISN"),
            "L/P": st.column_config.SelectboxColumn(
                "L/P", options=["", "L", "P"]
            ),
            "Agama": st.column_config.TextColumn("Agama"),
            "Umur": st.column_config.TextColumn("Umur"),
            "Pekerjaan Orang Tua": st.column_config.TextColumn(
                "Pekerjaan Orang Tua"
            ),
            "Tgl. Keluar": st.column_config.TextColumn("Tgl. Keluar"),
            "Tgl. Masuk": st.column_config.TextColumn("Tgl. Masuk"),
        },
        hide_index=True,
    )

    def generate_excel_mutasi_html(df_data, bulan, tahun):
        html = f"""
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
        <head>
            <meta charset="utf-8"/>
            <style>
                body {{ font-family: 'Times New Roman', Times, serif; }}
                table {{ border-collapse: collapse; width: 100%; }}
                th, td {{ border: 1px solid black; padding: 5px; text-align: center; }}
                .no-border {{ border: none !important; }}
                .header-title {{ font-weight: bold; font-size: 14pt; text-align: center; }}
            </style>
        </head>
        <body>
            <div class="header-title">MUTASI SISWA</div>
            <div class="header-title">{nama_sekolah.upper()}</div>
            <div class="header-title">TAHUN PELAJARAN {tahun_ajaran}</div>
            <br/><br/>
            <table class="no-border" style="width: auto; text-align: left;">
                <tr class="no-border"><td class="no-border" style="font-weight:bold;">Kelas</td><td class="no-border">: {kelas_input}</td></tr>
                <tr class="no-border"><td class="no-border" style="font-weight:bold;">Bulan</td><td class="no-border">: {nama_bulan[bulan-1]}</td></tr>
            </table>
            <br/>
            <table>
                <thead>
                    <tr style="font-weight:bold;">
                        <th style="width: 40px;">No</th>
                        <th style="width: 200px;">Nama Siswa</th>
                        <th style="width: 130px;">NIS / NISN</th>
                        <th style="width: 50px;">L/P</th>
                        <th style="width: 80px;">Agama</th>
                        <th style="width: 60px;">Umur</th>
                        <th style="width: 140px;">Pekerjaan Orang Tua</th>
                        <th style="width: 100px;">Tgl. Keluar</th>
                        <th style="width: 100px;">Tgl. Masuk</th>
                    </tr>
                </thead>
                <tbody>
        """

        for _, r in df_data.iterrows():
            html += f"""
                    <tr>
                        <td>{r['No']}</td>
                        <td align="left">{r['Nama Siswa']}</td>
                        <td style="mso-number-format:'\@';">{r['NIS / NISN']}</td>
                        <td>{r['L/P']}</td>
                        <td>{r['Agama']}</td>
                        <td>{r['Umur']}</td>
                        <td>{r['Pekerjaan Orang Tua']}</td>
                        <td>{r['Tgl. Keluar']}</td>
                        <td>{r['Tgl. Masuk']}</td>
                    </tr>
            """

        html += f"""
                </tbody>
            </table>
            <br/><br/>
            <table style="border:none; width:100%;">
                <tr style="border:none;">
                    <td style="border:none; width:60%;"></td>
                    <td style="border:none; text-align:center;">
                        {kota_lokasi}, 30 {nama_bulan[bulan-1]} {tahun}<br/>
                        Wali Kelas {kelas_input}<br/><br/><br/><br/>
                        <b><u>{nama_wali}</u></b><br/>
                        NIP. {nip_wali}
                    </td>
                </tr>
            </table>
        </body>
        </html>
        """
        return html.encode("utf-8")

    col_m_btn1, col_m_btn2 = st.columns([1, 1.5])

    with col_m_btn1:
        if st.button("💾 Simpan Data Mutasi", key="btn_save_mutasi_exact"):
            st.session_state["df_mutasi_exact"] = df_mutasi_edited
            st.success("✅ Data mutasi berhasil disimpan!")
            st.rerun()

    with col_m_btn2:
        excel_mutasi_bytes = generate_excel_mutasi_html(
            df_mutasi_edited, bln_mutasi, thn_mutasi
        )

        st.download_button(
            label="💾 UNDUH MUTASI SISWA (PERSIS SESUAI FOTO EXCEL)",
            data=excel_mutasi_bytes,
            file_name=f"Mutasi_Siswa_{nama_bulan[bln_mutasi-1]}_{thn_mutasi}.xls",
            mime="application/vnd.ms-excel",
            key="btn_download_mutasi_exact",
        )
