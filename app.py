import calendar
from datetime import datetime
import io
import pandas as pd
import streamlit as st

# ==========================================
# 1. KONFIGURASI HALAMAN
# ==========================================
st.set_page_config(
    page_title="SMP NEGERI 1 NANGA MAHAP - Rekapitulasi Absensi",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom CSS Styling
st.markdown(
    """
    <style>
    .header-container {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 20px;
        margin-bottom: 20px;
    }
    .title-header {
        text-align: center;
        font-weight: 800;
        font-size: 2.2rem;
        color: #212529;
        margin-bottom: 0px;
    }
    .subtitle-header {
        text-align: center;
        font-weight: 700;
        font-size: 1.2rem;
        color: #495057;
        margin-bottom: 20px;
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
# 2. HEADER SEKOLAH & LOGO (DENGAN UPLOAD LOGO)
# ==========================================
col_logo1, col_text, col_logo2 = st.columns([1, 4, 1])

with col_logo1:
    # Menggunakan gambar placeholder/logo default, atau bisa upload logo
    st.image(
        "https://cdn-icons-png.flaticon.com/512/2991/2991148.png", width=100
    )

with col_text:
    st.markdown(
        "<h1 class='title-header'>SMP NEGERI 1 NANGA MAHAP</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h3 class='subtitle-header'>REKAPITULASI ABSENSI & MUTASI SISWA - IX C (2025/2026)</h3>",
        unsafe_allow_html=True,
    )

with col_logo2:
    # Slot kosong/Simetris
    pass

st.divider()

# ==========================================
# 3. TAB NAVIGASI UTAMA (4 FITUR LENGKAP)
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
    st.caption("Ringkasan akumulasi absensi siswa per bulan.")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        bln_rekap = st.selectbox(
            "Pilih Bulan Rekap",
            options=list(range(1, 13)),
            format_func=lambda x: nama_bulan[x - 1],
            key="bln_tab1",
        )
    with col_b2:
        thn_rekap = st.number_input(
            "Pilih Tahun Rekap", value=datetime.now().year, key="thn_tab1"
        )

    # Menampilkan data siswa jika sudah diupload di Tab 3
    if "df_siswa" in st.session_state:
        df_display = st.session_state["df_siswa"].copy()
        if "Sakit (S)" not in df_display.columns:
            df_display["Sakit (S)"] = 0
            df_display["Izin (I)"] = 0
            df_display["Alpha (A)"] = 0
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    else:
        # Default Data Dummy jika belum upload
        data_bulanan = {
            "No": [1, 2, 3],
            "NIS/NISN": ["1001", "1002", "1003"],
            "Nama Siswa": ["Ahmad Fauzi", "Budi Santoso", "Citra Dewi"],
            "Sakit (S)": [1, 0, 2],
            "Izin (I)": [0, 1, 0],
            "Alpha (A)": [0, 0, 1],
            "Total Absen": [1, 1, 3],
            "Persentase Kehadiran": ["96%", "96%", "88%"],
        }
        st.dataframe(
            pd.DataFrame(data_bulanan),
            use_container_width=True,
            hide_index=True,
        )


# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN PER HARI
# ------------------------------------------
with tab2:
    st.subheader(
        "📅 Persentase Kehadiran Per Hari (Format Persis Gambar Upload)"
    )
    st.caption(
        "Isi nilai S, I, A atau tulis 'MINGGU' pada kolom Jumlah Siswa untuk menandai hari libur."
    )

    col1, col2, _ = st.columns([2, 2, 4])
    with col1:
        bulan_selected = st.selectbox(
            "Pilih Bulan",
            options=list(range(1, 13)),
            format_func=lambda x: f"{nama_bulan[x-1]} (Bulan {x})",
            index=datetime.now().month - 1,
            key="bln_tab2",
        )
    with col2:
        tahun_selected = st.number_input(
            "Pilih Tahun",
            min_value=2020,
            max_value=2035,
            value=datetime.now().year,
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

    # Fitur Tambah Tanggal Manual
    with st.expander("➕ Tambah Tanggal / Baris Baru Manual"):
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            tgl_baru = st.number_input(
                "Tanggal",
                min_value=1,
                max_value=31,
                value=total_hari + 1 if total_hari < 31 else 31,
            )
        with col_t2:
            jml_siswa_baru = st.text_input("Jumlah Siswa / Status", value="30")
        with col_t3:
            st.write("")
            st.write("")
            if st.button("Tambahkan Baris"):
                df_curr = st.session_state[session_key]
                new_row = pd.DataFrame(
                    [
                        {
                            "TGL": tgl_baru,
                            "Jumlah Siswa": jml_siswa_baru,
                            "S": 0,
                            "I": 0,
                            "A": 0,
                            "Jumlah": 0,
                            "Hadir %": (
                                "-"
                                if jml_siswa_baru.upper() == "MINGGU"
                                else "100%"
                            ),
                            "Tidak hadir %": (
                                "-"
                                if jml_siswa_baru.upper() == "MINGGU"
                                else "0%"
                            ),
                        }
                    ]
                )
                st.session_state[session_key] = pd.concat(
                    [df_curr, new_row], ignore_index=True
                )
                st.success(f"Tanggal {tgl_baru} berhasil ditambahkan!")
                st.rerun()

    # Fungsi Hitung Ulang Persentase
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
                    0.0, round(((tot_siswa - tot_absen) / tot_siswa) * 100, 1)
                )
                thp = min(100.0, round((tot_absen / tot_siswa) * 100, 1))
                h_list.append(f"{int(hp)}%" if hp.is_integer() else f"{hp}%")
                th_list.append(
                    f"{int(thp)}%" if thp.is_integer() else f"{thp}%"
                )

        df_copy["Jumlah"] = jml_list
        df_copy["Hadir %"] = h_list
        df_copy["Tidak hadir %"] = th_list
        return df_copy

    edited_df = st.data_editor(
        st.session_state[session_key],
        key=f"editor_{session_key}",
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "TGL": st.column_config.NumberColumn("↑ TGL"),
            "Jumlah Siswa": st.column_config.TextColumn("Jumlah Siswa"),
            "S": st.column_config.NumberColumn("S", min_value=0),
            "I": st.column_config.NumberColumn("I", min_value=0),
            "A": st.column_config.NumberColumn("A", min_value=0),
            "Jumlah": st.column_config.NumberColumn("Jumlah", disabled=True),
            "Hadir %": st.column_config.TextColumn("Hadir %", disabled=True),
            "Tidak hadir %": st.column_config.TextColumn(
                "Tidak hadir %", disabled=True
            ),
        },
        hide_index=True,
    )

    btn_col1, btn_col2 = st.columns([1, 1.5])
    with btn_col1:
        if st.button("💾 Simpan & Hitung Ulang Persentase", key="btn_save"):
            st.session_state[session_key] = hitung_ulang(edited_df)
            st.success("✅ Data berhasil dihitung ulang dan disimpan!")
            st.rerun()

    with btn_col2:
        df_export = hitung_ulang(edited_df)
        csv_data = df_export.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="💾 UNDUH FORMAT PERSENTASE KEHADIRAN HARIAN (EXACT SESUAI GAMBAR)",
            data=csv_data,
            file_name=f"Rekap_Kehadiran_Harian_{nama_bulan[bulan_selected-1]}_{tahun_selected}.csv",
            mime="text/csv",
            key="btn_download",
        )


# ------------------------------------------
# TAB 3: UPLOAD DATA SISWA
# ------------------------------------------
with tab4 if False else tab3:
    st.subheader("📂 Upload Data Siswa Kelas IX C")
    st.caption("Unggah file Excel (.xlsx / .xls) atau CSV daftar siswa Anda.")

    uploaded_file = st.file_uploader(
        "Pilih File Excel atau CSV", type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith(".csv"):
                df_uploaded = pd.read_csv(uploaded_file)
            else:
                df_uploaded = pd.read_excel(uploaded_file)

            st.session_state["df_siswa"] = df_uploaded
            st.success(
                f"✅ Berhasil mengunggah data {len(df_uploaded)} siswa!"
            )
            st.dataframe(df_uploaded, use_container_width=True)
        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

    if "df_siswa" in st.session_state:
        st.write("---")
        st.write("### Data Siswa Saat Ini:")
        st.dataframe(st.session_state["df_siswa"], use_container_width=True)


# ------------------------------------------
# TAB 4: MUTASI SISWA
# ------------------------------------------
with tab4:
    st.subheader("🔄 Mutasi Siswa")
    st.caption("Pencatatan data siswa masuk / keluar.")

    data_mutasi = {
        "No": [1],
        "Tanggal": ["2025-09-10"],
        "Nama Siswa": ["Rian Hidayat"],
        "Jenis Mutasi": ["Masuk"],
        "Keterangan": ["Pindahan dari SMPN 2"],
    }
    st.dataframe(
        pd.DataFrame(data_mutasi), use_container_width=True, hide_index=True
    )
