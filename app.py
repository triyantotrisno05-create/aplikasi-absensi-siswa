import calendar
from datetime import datetime
import io
import pandas as pd
import streamlit as st

# Config Halaman
st.set_page_config(
    page_title="SMP NEGERI 1 NANGA MAHAP - Rekapitulasi Absensi",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom Styling (Menyamakan tampilan dengan screenshot)
st.markdown(
    """
    <style>
    .main-title {
        text-align: center;
        font-weight: 800;
        font-size: 2.2rem;
        color: #1a1a1a;
        margin-bottom: 5px;
    }
    .sub-title {
        text-align: center;
        font-weight: 700;
        font-size: 1.25rem;
        color: #333333;
        margin-bottom: 25px;
    }
    
    /* Style Tombol Merah Sesuai Gambar */
    div.stButton > button:first-child {
        background-color: #ff4d4f !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 12px 20px !important;
        font-size: 14px !important;
        width: 100%;
    }
    div.stButton > button:first-child:hover {
        background-color: #e03e3f !important;
        color: white !important;
    }

    div.stDownloadButton > button:first-child {
        background-color: #ff4d4f !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 12px 20px !important;
        font-size: 14px !important;
        width: 100%;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #e03e3f !important;
        color: white !important;
    }

    /* Style Tab */
    .stTabs [data-baseweb="tab-list"] {
        gap: 20px;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: 600;
        font-size: 15px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# Judul Sekolah
st.markdown(
    "<div class='main-title'>SMP NEGERI 1 NANGA MAHAP</div>",
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='sub-title'>REKAPITULASI ABSENSI & MUTASI SISWA - IX C (2025/2026)</div>",
    unsafe_allow_html=True,
)

# Tab Menu Aplikasi
tab1, tab2, tab3 = st.tabs(
    ["📊 Rekap Absensi Bulanan", "📅 Persentase Kehadiran Per Hari", "🔄 Mutasi Siswa"]
)

with tab2:
    st.subheader("📅 Persentase Kehadiran Per Hari (Format Persis Gambar Upload)")
    st.caption(
        "Isi nilai S, I, A atau tulis 'MINGGU' pada kolom Jumlah Siswa untuk menandai hari libur."
    )

    # Filter Bulan dan Tahun
    col_filter1, col_filter2, col_filter3 = st.columns([2, 2, 3])

    nama_bulan = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]

    with col_filter1:
        bulan_selected = st.selectbox(
            "Pilih Bulan",
            options=list(range(1, 13)),
            format_func=lambda x: f"{nama_bulan[x-1]} (Bulan {x})",
            index=datetime.now().month - 1,
        )

    with col_filter2:
        tahun_selected = st.number_input(
            "Pilih Tahun",
            min_value=2020,
            max_value=2035,
            value=datetime.now().year,
        )

    # Menghitung Jumlah Hari Otomatis Dalam Bulan (28, 29, 30, atau 31)
    _, total_hari = calendar.monthrange(tahun_selected, bulan_selected)

    # Inisialisasi Data Session State
    session_key = f"df_absensi_{bulan_selected}_{tahun_selected}"

    if session_key not in st.session_state:
        data_list = []
        for tgl in range(1, total_hari + 1):
            is_minggu = (
                calendar.weekday(tahun_selected, bulan_selected, tgl) == 6
            )
            data_list.append(
                {
                    "TGL": int(tgl),
                    "Jumlah Siswa": "MINGGU" if is_minggu else "30",
                    "S": 0,
                    "I": 0,
                    "A": 0,
                    "Jumlah": 0,
                    "Hadir %": "-" if is_minggu else "100%",
                    "Tidak hadir %": "-" if is_minggu else "0%",
                }
            )
        st.session_state[session_key] = pd.DataFrame(data_list)

    # Fungsi Rekalkulasi Otomatis Persentase Kehadiran
    def kalkulasi_persentase(df):
        df_copy = df.copy()
        jumlah_list = []
        hadir_pct_list = []
        tidak_hadir_pct_list = []

        for idx, row in df_copy.iterrows():
            jml_siswa_str = str(row["Jumlah Siswa"]).strip()
            
            # Hitung total S, I, A
            try:
                s = int(row["S"])
            except:
                s = 0
            try:
                i = int(row["I"])
            except:
                i = 0
            try:
                a = int(row["A"])
            except:
                a = 0

            tot_absen = s + i + a
            jumlah_list.append(tot_absen)

            # Jika 'MINGGU' atau Hari Libur
            if (
                jml_siswa_str.upper() == "MINGGU"
                or not jml_siswa_str.isdigit()
                or int(jml_siswa_str) == 0
            ):
                hadir_pct_list.append("-")
                tidak_hadir_pct_list.append("-")
            else:
                tot_siswa = int(jml_siswa_str)
                h_pct = max(
                    0.0, round(((tot_siswa - tot_absen) / tot_siswa) * 100, 1)
                )
                th_pct = min(100.0, round((tot_absen / tot_siswa) * 100, 1))

                hadir_pct_list.append(
                    f"{int(h_pct)}%" if h_pct.is_integer() else f"{h_pct}%"
                )
                tidak_hadir_pct_list.append(
                    f"{int(th_pct)}%" if th_pct.is_integer() else f"{th_pct}%"
                )

        df_copy["Jumlah"] = jumlah_list
        df_copy["Hadir %"] = hadir_pct_list
        df_copy["Tidak hadir %"] = tidak_hadir_pct_list
        return df_copy

    st.write("---")

    # Fitur Tambah Baris/Tanggal Baru
    with st.expander("➕ **Tambah Tanggal / Baris Baru Manual**", expanded=False):
        col_add1, col_add2, col_add3 = st.columns([2, 3, 2])
        with col_add1:
            next_tgl = len(st.session_state[session_key]) + 1
            tgl_baru = st.number_input("Tanggal (TGL)", min_value=1, value=next_tgl)
        with col_add2:
            jml_siswa_baru = st.text_input("Jumlah Siswa / Tulis 'MINGGU'", value="30")
        with col_add3:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("➕ Tambah Baris"):
                new_row = pd.DataFrame([{
                    "TGL": int(tgl_baru),
                    "Jumlah Siswa": jml_siswa_baru,
                    "S": 0,
                    "I": 0,
                    "A": 0,
                    "Jumlah": 0,
                    "Hadir %": "100%",
                    "Tidak hadir %": "0%"
                }])
                st.session_state[session_key] = pd.concat([st.session_state[session_key], new_row], ignore_index=True)
                st.session_state[session_key] = kalkulasi_persentase(st.session_state[session_key])
                st.success(f"Tanggal {tgl_baru} berhasil ditambahkan!")
                st.rerun()

    # Tabel Interaktif (Bisa Didefinisikan & Diedit Langsung)
    df_tampil = st.session_state[session_key]

    edited_df = st.data_editor(
        df_tampil,
        key=f"editor_{session_key}",
        num_rows="dynamic", # Memungkinkan edit, hapus, dan tambah baris langsung dari tabel
        use_container_width=True,
        column_config={
            "TGL": st.column_config.NumberColumn(
                "↑ TGL", required=True, width="small"
            ),
            "Jumlah Siswa": st.column_config.TextColumn(
                "Jumlah Siswa", required=True, width="medium"
            ),
            "S": st.column_config.NumberColumn("S", min_value=0, width="small", default=0),
            "I": st.column_config.NumberColumn("I", min_value=0, width="small", default=0),
            "A": st.column_config.NumberColumn("A", min_value=0, width="small", default=0),
            "Jumlah": st.column_config.NumberColumn("Jumlah", disabled=True),
            "Hadir %": st.column_config.TextColumn("Hadir %", disabled=True),
            "Tidak hadir %": st.column_config.TextColumn(
                "Tidak hadir %", disabled=True
            ),
        },
        hide_index=True,
    )

    # Tombol Simpan & Unduh (Persis Tombol Merah Aplikasi Anda)
    col_btn1, col_btn2 = st.columns([1, 1.5])

    with col_btn1:
        if st.button("💾 Simpan & Hitung Ulang Persentase"):
            df_updated = kalkulasi_persentase(edited_df)
            st.session_state[session_key] = df_updated
            st.success("✅ Data absensi berhasil disimpan dan dihitung ulang!")
            st.rerun()

    with col_btn2:
        df_for_download = kalkulasi_persentase(edited_df)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df_for_download.to_excel(
                writer,
                sheet_name=f"Rekap_{nama_bulan[bulan_selected-1]}",
                index=False,
            )

        st.download_button(
            label="💾 UNDUH FORMAT PERSENTASE KEHADIRAN HARIAN (EXACT SESUAI GAMBAR)",
            data=buffer.getvalue(),
            file_name=f"Rekap_Kehadiran_Harian_{nama_bulan[bulan_selected-1]}_{tahun_selected}.xlsx",
            mime="application/vnd.ms-excel",
        )

with tab1:
    st.info("Fitur Rekap Absensi Bulanan")

with tab3:
    st.info("Fitur Mutasi Siswa")
