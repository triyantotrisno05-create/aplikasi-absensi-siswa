import calendar
import json
import os
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
    .title-header { text-align: center; font-weight: 800; font-size: 2.0rem; color: #212529; margin-bottom: 0px; }
    .subtitle-header { text-align: center; font-weight: 700; font-size: 1.1rem; color: #495057; margin-bottom: 15px; }
    div.stButton > button:first-child, div.stDownloadButton > button:first-child {
        background-color: #ff4d4f !important; color: white !important; font-weight: bold !important;
        border-radius: 6px !important; border: none !important; padding: 10px 20px !important;
        font-size: 14px !important; width: 100%; margin-top: 5px;
    }
    div.stButton > button:first-child:hover, div.stDownloadButton > button:first-child:hover {
        background-color: #ff2a2d !important; color: white !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. SISTEM PENYIMPANAN DATA PER KELAS (JSON)
# ==========================================
DATA_DIR = "data_kelas"
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)


def get_file_path(kelas_nama):
    clean_name = "".join(
        c for c in kelas_nama if c.isalnum() or c in (" ", "_", "-")
    ).rstrip()
    return os.path.join(DATA_DIR, f"data_{clean_name}.json")


def simpan_data_kelas(kelas_nama, data_dict):
    filepath = get_file_path(kelas_nama)
    with open(filepath, "w") as f:
        json.dump(data_dict, f, indent=4)


def muat_data_kelas(kelas_nama):
    filepath = get_file_path(kelas_nama)
    if os.path.exists(filepath):
        with open(filepath, "r") as f:
            return json.load(f)
    return None


# ==========================================
# 3. PENGATURAN IDENTITAS SEKOLAH (SIDEBAR)
# ==========================================
with st.sidebar:
    st.header("⚙️ Pengaturan Header & Data")

    nama_sekolah = st.text_input(
        "Nama Sekolah", value="SMP NEGERI 1 NANGA MAHAP"
    )
    kelas_input = st.text_input("Kelas Anda", value="IX C")
    tahun_ajaran = st.text_input("Tahun Pelajaran", value="2025/2026")
    nama_wali = st.text_input("Nama Wali Kelas", value="Trivanto trisno, S.Pd")
    nip_wali = st.text_input("NIP Wali Kelas", value="199305202024211000")
    kota_lokasi = st.text_input("Kota / Kecamatan", value="Nanga Mahap")

    st.info(f"📌 Anda sedang mengedit data untuk **Kelas: {kelas_input}**")

# Muat data tersimpan untuk kelas ini jika ada
data_tersimpan = muat_data_kelas(kelas_input)

# ==========================================
# 4. TAMPILAN HEADER UTAMA
# ==========================================
col_logo1, col_text, col_logo2 = st.columns([1, 4, 1])

with col_logo1:
    st.image("https://cdn-icons-png.flaticon.com/512/2991/2991148.png", width=90)

with col_text:
    st.markdown(
        f"<h1 class='title-header'>{nama_sekolah.upper()}</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<h3 class='subtitle-header'>REKAPITULASI ABSENSI & MUTASI SISWA - KELAS {kelas_input.upper()} ({tahun_ajaran})</h3>",
        unsafe_allow_html=True,
    )

st.divider()

# ==========================================
# 5. TAB NAVIGASI UTAMA
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
    st.subheader(f"📊 Rekap Absensi Bulanan (Kelas {kelas_input})")

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
            "Pilih Tahun Rekap", value=datetime.now().year, key="thn_tab1"
        )
    with col_b3:
        hbe_input = st.number_input(
            "Hari Belajar Efektif (HBE)", min_value=1, max_value=31, value=25
        )

    # Inisialisasi Data Siswa
    if (
        data_tersimpan
        and "df_siswa" in data_tersimpan
        and len(data_tersimpan["df_siswa"]) > 0
    ):
        df_base = pd.DataFrame(data_tersimpan["df_siswa"])
    elif "df_siswa" in st.session_state and not st.session_state[
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

    session_key_rekap = f"df_rekap_{kelas_input}_{bln_rekap}_{thn_rekap}"

    if session_key_rekap not in st.session_state:
        df_init = df_base.copy()
        df_init["HBE"] = hbe_input
        if "S" not in df_init.columns:
            df_init["S"] = 0
        if "I" not in df_init.columns:
            df_init["I"] = 0
        if "A" not in df_init.columns:
            df_init["A"] = 0
        st.session_state[session_key_rekap] = df_init

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
        hide_index=True,
    )

    df_final_rekap = hitung_rekap_bulanan(df_rekap_edited, hbe_input)

    if st.button(f"💾 Simpan Data Kelas {kelas_input}", key="save_rekap"):
        p_data = data_tersimpan if data_tersimpan else {}
        p_data["df_siswa"] = df_base.to_dict(orient="records")
        p_data[f"rekap_{bln_rekap}_{thn_rekap}"] = df_rekap_edited.to_dict(
            orient="records"
        )
        simpan_data_kelas(kelas_input, p_data)
        st.success(f"✅ Data Rekapitulasi Kelas {kelas_input} Berhasil Disimpan!")

# ------------------------------------------
# TAB 2: PERSENTASE KEHADIRAN PER HARI
# ------------------------------------------
with tab2:
    st.subheader(
        f"📅 Persentase Kehadiran Per Hari (Kelas {kelas_input} - Format Persis Foto)"
    )

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
    session_key = f"df_harian_{kelas_input}_{bulan_selected}_{tahun_selected}"

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

    if st.button(
        f"💾 Simpan Kehadiran Harian (Kelas {kelas_input})", key="btn_save_harian"
    ):
        p_data = data_tersimpan if data_tersimpan else {}
        p_data[f"harian_{bulan_selected}_{tahun_selected}"] = (
            edited_df.to_dict(orient="records")
        )
        simpan_data_kelas(kelas_input, p_data)
        st.success(f"✅ Data Harian Kelas {kelas_input} Berhasil Disimpan!")

# ------------------------------------------
# TAB 3: UPLOAD DATA SISWA
# ------------------------------------------
with tab3:
    st.subheader(f"📂 Upload Data Siswa Kelas {kelas_input}")

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

            # Simpan ke file kelas
            p_data = data_tersimpan if data_tersimpan else {}
            p_data["df_siswa"] = df_uploaded.to_dict(orient="records")
            simpan_data_kelas(kelas_input, p_data)

            st.success(
                f"✅ BERHASIL! Data {len(df_uploaded)} siswa untuk Kelas {kelas_input} telah tersimpan!"
            )
            st.rerun()

        except Exception as e:
            st.error(f"Gagal membaca file: {e}")

# ------------------------------------------
# TAB 4: MUTASI SISWA
# ------------------------------------------
with tab4:
    st.subheader(f"🔄 Data Mutasi Siswa Kelas {kelas_input}")

    if "df_mutasi" not in st.session_state:
        st.session_state["df_mutasi"] = pd.DataFrame(
            [
                {
                    "No": "1",
                    "Nama Siswa": "",
                    "NIS / NISN": "",
                    "L/P": "",
                    "Agama": "",
                    "Umur": "",
                    "Pekerjaan Orang Tua": "",
                    "Tgl. Keluar": "",
                    "Tgl. Masuk": "",
                }
            ]
        )

    df_mutasi_edited = st.data_editor(
        st.session_state["df_mutasi"],
        key="editor_mutasi",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
    )

    if st.button(
        f"💾 Simpan Mutasi Siswa (Kelas {kelas_input})", key="btn_save_mutasi"
    ):
        p_data = data_tersimpan if data_tersimpan else {}
        p_data["df_mutasi"] = df_mutasi_edited.to_dict(orient="records")
        simpan_data_kelas(kelas_input, p_data)
        st.success(f"✅ Data Mutasi Kelas {kelas_input} Berhasil Disimpan!")
