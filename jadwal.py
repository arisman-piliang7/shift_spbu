import streamlit as st
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string
from datetime import date, time, timedelta
from collections import OrderedDict
from io import BytesIO

# =========================================================
# 1. KONFIGURASI HALAMAN & TEMA
# =========================================================

st.set_page_config(
    page_title="Penjadwalan Operator SPBU",
    page_icon="📅",
    layout="wide",
)

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.block-container { padding-top: 2rem; max-width: 1200px; }

/* ---------- Hero ---------- */
.hero {
    background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 55%, #db2777 120%);
    border-radius: 22px;
    padding: 2rem 2.2rem;
    color: #fff;
    box-shadow: 0 18px 40px -18px rgba(79,70,229,.55);
    margin-bottom: 1.4rem;
    position: relative;
    overflow: hidden;
}
.hero::after {
    content: "";
    position: absolute; right: -60px; top: -60px;
    width: 220px; height: 220px; border-radius: 50%;
    background: rgba(255,255,255,.12);
}
.hero h1 { margin: 0 0 .35rem 0; font-size: 1.85rem; font-weight: 700; color: #fff; }
.hero p  { margin: 0; opacity: .9; font-size: .98rem; max-width: 640px; }

/* ---------- Langkah ---------- */
.step {
    display: flex; align-items: center; gap: .7rem;
    margin: 1.6rem 0 .6rem 0;
}
.step .num {
    width: 32px; height: 32px; border-radius: 10px;
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    color: #fff; font-weight: 700; font-size: .95rem;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 6px 14px -6px rgba(99,102,241,.8);
}
.step .ttl { font-size: 1.2rem; font-weight: 600; }
.step .sub { font-size: .85rem; opacity: .65; margin-left: .2rem; }

/* ---------- Chip ---------- */
.chips { display: flex; flex-wrap: wrap; gap: .45rem; margin: .4rem 0 .8rem 0; }
.chip {
    background: rgba(99,102,241,.12);
    color: #6366f1;
    border: 1px solid rgba(99,102,241,.25);
    padding: .22rem .7rem; border-radius: 999px;
    font-size: .8rem; font-weight: 500;
}
.chip.warn { background: rgba(245,158,11,.12); color: #d97706; border-color: rgba(245,158,11,.3); }
.chip.off  { background: rgba(148,163,184,.15); color: #64748b; border-color: rgba(148,163,184,.3); text-decoration: line-through; }

/* ---------- Metric ---------- */
div[data-testid="stMetric"] {
    background: rgba(99,102,241,.07);
    border: 1px solid rgba(99,102,241,.18);
    border-radius: 16px;
    padding: .9rem 1.1rem;
}
div[data-testid="stMetricValue"] { font-weight: 700; }

/* ---------- Tombol ---------- */
.stButton > button, .stDownloadButton > button {
    border-radius: 14px; font-weight: 600; padding: .7rem 1rem;
    transition: transform .15s ease, box-shadow .15s ease;
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    border: none; color: #fff;
    box-shadow: 0 10px 22px -10px rgba(99,102,241,.9);
}
.stButton > button:hover, .stDownloadButton > button:hover { transform: translateY(-1px); }

/* ---------- Lain-lain ---------- */
section[data-testid="stSidebar"] { border-right: 1px solid rgba(148,163,184,.25); }
section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { font-size: 1.02rem; }
div[data-testid="stExpander"] { border-radius: 14px; }
div[data-testid="stFileUploader"] section { border-radius: 16px; border: 2px dashed rgba(99,102,241,.4); }
</style>
"""

st.markdown(CSS, unsafe_allow_html=True)

st.markdown(
    """
    <div class="hero">
        <h1>📅 Penjadwalan Pelatihan Operator SPBU</h1>
        <p>Susun jadwal pelatihan otomatis berdasarkan nomor SPBU. Atur sendiri
        hari yang dikecualikan, jumlah shift, dan jumlah operator per shift,
        lalu unduh hasilnya dalam format Excel dengan merge cells rapi.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


def judul_langkah(nomor, judul, sub=""):
    st.markdown(
        f"""
        <div class="step">
            <div class="num">{nomor}</div>
            <div><span class="ttl">{judul}</span>
            <span class="sub">{sub}</span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# 2. KONSTANTA
# =========================================================

NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]

BARIS_AWAL = 2

# Nama & jam bawaan untuk shift (bisa diubah di sidebar)
SHIFT_BAWAAN = [
    ("Pagi", time(8, 0), time(11, 0)),
    ("Siang", time(13, 0), time(16, 0)),
    ("Sore", time(16, 30), time(19, 30)),
    ("Malam", time(19, 30), time(22, 30)),
    ("Shift 5", time(6, 0), time(8, 0)),
    ("Shift 6", time(22, 30), time(23, 59)),
]

MODE_ISI = {
    "Per tanggal (semua shift hari itu terisi dulu)": "tanggal",
    "Rotasi tanggal (tiap kelompok pindah hari)": "rotasi",
}


def fmt_jam(t):
    return t.strftime("%H.%M")


def label_shift(nama, mulai, selesai):
    return f"Sesi {nama} ({fmt_jam(mulai)}-{fmt_jam(selesai)})"


# =========================================================
# 3. FUNGSI TANGGAL
# =========================================================


def daftar_tanggal(mulai, maksimum, hari_dikecualikan, tanggal_khusus):
    """Semua tanggal dalam rentang, tanpa hari/tanggal yang dikecualikan."""

    hasil = []
    khusus = set(tanggal_khusus)
    hari = mulai

    while hari <= maksimum:

        if hari.weekday() not in hari_dikecualikan and hari not in khusus:
            hasil.append(hari)

        hari += timedelta(days=1)

    return hasil


def semua_tanggal(mulai, maksimum):

    return [mulai + timedelta(days=i) for i in range((maksimum - mulai).days + 1)]


def tentukan_slot(index, tanggal_list, jumlah_shift, paralel, mode):
    """Mengembalikan (tanggal, index_shift) untuk kelompok ke-index."""

    slot = index // paralel
    total_hari = len(tanggal_list)

    if mode == "rotasi":
        idx_tanggal = slot % total_hari
        idx_shift = (slot // total_hari) % jumlah_shift
    else:
        idx_tanggal = (slot // jumlah_shift) % total_hari
        idx_shift = slot % jumlah_shift

    return tanggal_list[idx_tanggal], idx_shift


# =========================================================
# 4. FUNGSI MEMBUAT KELOMPOK
# =========================================================


def buat_kelompok_operator(ws, kolom_spbu, kolom_operator):

    kelompok = OrderedDict()

    for row in range(BARIS_AWAL, ws.max_row + 1):

        nomor_spbu = ws.cell(row, kolom_spbu).value
        nama_operator = ws.cell(row, kolom_operator).value

        if nomor_spbu is None or nama_operator is None:
            continue

        if not str(nama_operator).strip():
            continue

        nomor_spbu = str(nomor_spbu).strip()

        kelompok.setdefault(nomor_spbu, []).append(
            {"row": row, "spbu": nomor_spbu, "operator": str(nama_operator).strip()}
        )

    return kelompok


def buat_grup(kelompok, ukuran, izinkan_sisa):
    """
    Operator dikelompokkan per SPBU menjadi grup berisi `ukuran` orang.
    Operator sisa dibawa ke SPBU berikutnya.
    """

    grup = []
    sisa = []

    for operators in kelompok.values():

        daftar = sisa + operators
        sisa = []

        i = 0

        while i + ukuran <= len(daftar):
            grup.append(daftar[i : i + ukuran])
            i += ukuran

        sisa = daftar[i:]

    if sisa:

        if izinkan_sisa:
            grup.append(sisa)

        else:
            nama = ", ".join(o["operator"] for o in sisa)
            raise ValueError(
                f"Operator {nama} belum memiliki kelompok penuh "
                f"({ukuran} operator per shift). Total operator harus "
                f"kelipatan {ukuran}, atau aktifkan opsi "
                "'Izinkan kelompok terakhir tidak penuh'."
            )

    return grup


# =========================================================
# 5. FUNGSI MEMPROSES JADWAL
# =========================================================


def proses_penjadwalan(
    file_bytes,
    nama_sheet,
    tanggal_list,
    shifts,
    ukuran_grup,
    paralel,
    mode,
    izinkan_sisa,
    kolom,
):

    kolom_spbu, kolom_operator, kolom_jadwal, kolom_tanggal = kolom

    wb = load_workbook(BytesIO(file_bytes))

    if nama_sheet not in wb.sheetnames:
        raise ValueError("Sheet tidak ditemukan.")

    ws = wb[nama_sheet]

    kelompok = buat_kelompok_operator(ws, kolom_spbu, kolom_operator)

    if not kelompok:
        raise ValueError("Tidak ditemukan data operator pada kolom SPBU dan Operator.")

    grup_list = buat_grup(kelompok, ukuran_grup, izinkan_sisa)

    if not grup_list:
        raise ValueError("Tidak ada kelompok operator yang terbentuk.")

    # Merge vertikal membutuhkan baris berurutan
    for grup in grup_list:

        baris = sorted(o["row"] for o in grup)

        if any(b2 - b1 != 1 for b1, b2 in zip(baris, baris[1:])):

            nama = ", ".join(o["operator"] for o in grup)
            raise ValueError(
                f"Kelompok ({nama}) berada pada baris {baris} "
                "yang tidak berurutan. Susun ulang data operator "
                "agar tiap kelompok berada pada baris bersebelahan."
            )

    # Lepas merge lama pada kolom jadwal & tanggal
    kolom_min = min(kolom_jadwal, kolom_tanggal)
    kolom_max = max(kolom_jadwal, kolom_tanggal)

    for merged in list(ws.merged_cells.ranges):

        if (
            merged.min_col <= kolom_max
            and merged.max_col >= kolom_min
            and merged.max_row >= BARIS_AWAL
        ):
            ws.unmerge_cells(str(merged))

    # Bersihkan nilai lama
    for row in range(BARIS_AWAL, ws.max_row + 1):
        ws.cell(row, kolom_jadwal).value = None
        ws.cell(row, kolom_tanggal).value = None

    # Buat jadwal
    hasil = []

    for index, grup in enumerate(grup_list):

        tanggal, idx_shift = tentukan_slot(
            index, tanggal_list, len(shifts), paralel, mode
        )

        sesi = shifts[idx_shift]

        row_awal = min(o["row"] for o in grup)
        row_akhir = max(o["row"] for o in grup)

        ws.cell(row_awal, kolom_jadwal).value = sesi
        ws.cell(row_awal, kolom_tanggal).value = tanggal
        ws.cell(row_awal, kolom_tanggal).number_format = "dd.mm.yyyy"

        if row_akhir > row_awal:

            for kol in (kolom_jadwal, kolom_tanggal):
                ws.merge_cells(
                    start_row=row_awal,
                    start_column=kol,
                    end_row=row_akhir,
                    end_column=kol,
                )

        ws.cell(row_awal, kolom_jadwal).alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        ws.cell(row_awal, kolom_tanggal).alignment = Alignment(
            horizontal="center", vertical="center"
        )

        for operator in grup:

            hasil.append(
                {
                    "Kelompok": index + 1,
                    "Nomor SPBU": operator["spbu"],
                    "Nama Operator": operator["operator"],
                    "Hari": NAMA_HARI[tanggal.weekday()],
                    "Tanggal": tanggal.strftime("%d.%m.%Y"),
                    "Jadwal": sesi,
                    "Baris Excel": operator["row"],
                }
            )

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue(), hasil, len(kelompok), len(grup_list)


# =========================================================
# 6. SIDEBAR PENGATURAN
# =========================================================

with st.sidebar:

    st.markdown("## ⚙️ Pengaturan Jadwal")

    # ---------- Rentang tanggal ----------
    st.markdown("### 🗓️ Rentang Tanggal")

    tanggal_mulai = st.date_input("Tanggal mulai", value=date(2026, 11, 9))
    tanggal_maksimum = st.date_input("Tanggal maksimum", value=date(2026, 11, 27))

    # ---------- Hari dikecualikan ----------
    st.markdown("### 🚫 Hari Dikecualikan")

    hari_dikecualikan_nama = st.multiselect(
        "Hari yang tidak ada pelatihan",
        options=NAMA_HARI,
        default=["Sabtu", "Minggu"],
        help="Pilih hari apa saja yang dilewati. Kosongkan jika semua hari aktif.",
    )
    hari_dikecualikan = {NAMA_HARI.index(h) for h in hari_dikecualikan_nama}

    tanggal_khusus = []

    if tanggal_maksimum >= tanggal_mulai:

        kandidat = [
            t
            for t in semua_tanggal(tanggal_mulai, tanggal_maksimum)
            if t.weekday() not in hari_dikecualikan
        ]

        tanggal_khusus = st.multiselect(
            "Tanggal libur khusus (opsional)",
            options=kandidat,
            format_func=lambda d: f"{NAMA_HARI[d.weekday()]}, {d.strftime('%d.%m.%Y')}",
            help="Contoh: tanggal merah atau cuti bersama.",
        )

    # ---------- Shift ----------
    st.markdown("### ⏱️ Shift & Kapasitas")

    jumlah_shift = st.number_input(
        "Jumlah shift per hari", min_value=1, max_value=6, value=2, step=1
    )

    ukuran_grup = st.number_input(
        "Operator per shift (per kelompok)",
        min_value=1,
        max_value=10,
        value=2,
        step=1,
        help="Jumlah operator yang digabung dalam satu sesi pelatihan.",
    )

    paralel = st.number_input(
        "Kelompok paralel per shift",
        min_value=1,
        max_value=20,
        value=1,
        step=1,
        help="Berapa kelompok yang boleh berlatih bersamaan pada shift yang sama.",
    )

    shifts = []

    for i in range(int(jumlah_shift)):

        nama_b, mulai_b, selesai_b = SHIFT_BAWAAN[i]

        with st.expander(f"Shift {i + 1}: {nama_b}", expanded=(i < 2)):

            nama = st.text_input("Nama shift", value=nama_b, key=f"shift_nama_{i}")

            c1, c2 = st.columns(2)

            mulai = c1.time_input("Mulai", value=mulai_b, key=f"shift_mulai_{i}")
            selesai = c2.time_input(
                "Selesai", value=selesai_b, key=f"shift_selesai_{i}"
            )

            shifts.append(label_shift(nama.strip() or f"Shift {i + 1}", mulai, selesai))

    # ---------- Pola pengisian ----------
    st.markdown("### 🔁 Pola Pengisian")

    mode_label = st.radio("Urutan pengisian jadwal", list(MODE_ISI.keys()))
    mode = MODE_ISI[mode_label]

    izinkan_sisa = st.toggle(
        "Izinkan kelompok terakhir tidak penuh",
        value=False,
        help="Jika total operator tidak habis dibagi jumlah operator per shift.",
    )

    # ---------- Kolom ----------
    with st.expander("🧩 Pengaturan kolom Excel"):

        huruf_spbu = st.text_input("Kolom Nomor SPBU", value="C").strip().upper()
        huruf_operator = st.text_input("Kolom Nama Operator", value="D").strip().upper()
        huruf_jadwal = st.text_input("Kolom Jadwal (hasil)", value="H").strip().upper()
        huruf_tanggal = (
            st.text_input("Kolom Tanggal (hasil)", value="I").strip().upper()
        )


# =========================================================
# 7. RINGKASAN PENGATURAN
# =========================================================

tanggal_list = []

if tanggal_maksimum >= tanggal_mulai:

    tanggal_list = daftar_tanggal(
        tanggal_mulai, tanggal_maksimum, hari_dikecualikan, tanggal_khusus
    )

kapasitas = len(tanggal_list) * len(shifts) * int(paralel)

judul_langkah("1", "Ringkasan Pengaturan", "pastikan sudah sesuai")

m1, m2, m3, m4 = st.columns(4)

m1.metric("Hari aktif", len(tanggal_list))
m2.metric("Shift per hari", len(shifts))
m3.metric("Operator per shift", int(ukuran_grup))
m4.metric("Kapasitas kelompok", kapasitas)

if tanggal_maksimum < tanggal_mulai:

    st.error("Tanggal maksimum harus sama dengan atau setelah tanggal mulai.")

elif not tanggal_list:

    st.error("Tidak ada hari aktif pada rentang ini. Kurangi hari yang dikecualikan.")

else:

    chips = "".join(
        f'<span class="chip">{NAMA_HARI[d.weekday()][:3]} {d.strftime("%d.%m")}</span>'
        for d in tanggal_list
    )
    chips += "".join(
        f'<span class="chip off">{h}</span>' for h in hari_dikecualikan_nama
    )

    st.markdown(f'<div class="chips">{chips}</div>', unsafe_allow_html=True)

    st.caption("Sesi: " + "  •  ".join(shifts) + f"   |   Pola: {mode_label.lower()}")

# =========================================================
# 8. UPLOAD EXCEL
# =========================================================

judul_langkah("2", "Upload Data Operator", "format .xlsx")

uploaded_file = st.file_uploader(
    "Tarik file ke sini atau klik untuk memilih", type=["xlsx"]
)

if uploaded_file is not None:

    file_bytes = uploaded_file.getvalue()

    try:

        wb_preview = load_workbook(BytesIO(file_bytes), read_only=True)
        daftar_sheet = wb_preview.sheetnames
        wb_preview.close()

        st.success(f"File berhasil dibaca: {uploaded_file.name}")

        judul_langkah("3", "Pilih Sheet & Generate")

        index_default = (
            daftar_sheet.index("dataoperator(2)")
            if "dataoperator(2)" in daftar_sheet
            else 0
        )

        nama_sheet = st.selectbox(
            "Sheet yang akan diproses", daftar_sheet, index=index_default
        )

        with st.expander("ℹ️ Lihat aturan penjadwalan"):

            st.markdown(f"""
                1. Operator dikelompokkan berdasarkan nomor SPBU.
                2. Setiap kelompok berisi **{int(ukuran_grup)} operator** dan mendapat satu sesi yang sama.
                3. Operator sisa dibawa ke SPBU berikutnya.
                4. Hari yang dikecualikan dilewati: **{', '.join(hari_dikecualikan_nama) or 'tidak ada'}**.
                5. Tersedia **{len(shifts)} shift** per hari, **{int(paralel)} kelompok** per shift.
                6. Jika kapasitas terlewati, jadwal berputar kembali ke awal rentang.
                7. Kolom jadwal dan tanggal di-merge untuk setiap kelompok.
                """)

        if st.button(
            "🚀 Generate Jadwal Otomatis",
            type="primary",
            use_container_width=True,
            disabled=not tanggal_list,
        ):

            try:

                kolom = (
                    column_index_from_string(huruf_spbu),
                    column_index_from_string(huruf_operator),
                    column_index_from_string(huruf_jadwal),
                    column_index_from_string(huruf_tanggal),
                )

                with st.spinner("Sedang membuat jadwal operator..."):

                    output, hasil, jumlah_spbu, jumlah_grup = proses_penjadwalan(
                        file_bytes,
                        nama_sheet,
                        tanggal_list,
                        shifts,
                        int(ukuran_grup),
                        int(paralel),
                        mode,
                        izinkan_sisa,
                        kolom,
                    )

                st.session_state["output"] = output
                st.session_state["hasil"] = hasil
                st.session_state["jumlah_spbu"] = jumlah_spbu
                st.session_state["jumlah_grup"] = jumlah_grup
                st.session_state["shifts"] = list(shifts)
                st.session_state["kapasitas"] = kapasitas

                st.success("Jadwal berhasil dibuat! 🎉")

            except Exception as e:

                st.error(f"Gagal membuat jadwal: {e}")

    except Exception as e:

        st.error(f"Gagal membaca file: {e}")


# =========================================================
# 9. HASIL
# =========================================================

if "hasil" in st.session_state:

    judul_langkah("4", "Hasil Penjadwalan")

    hasil = st.session_state["hasil"]
    df = pd.DataFrame(hasil)

    c1, c2, c3, c4 = st.columns(4)

    c1.metric("Jumlah SPBU", st.session_state["jumlah_spbu"])
    c2.metric("Jumlah Kelompok", st.session_state["jumlah_grup"])
    c3.metric("Jumlah Operator", len(df))
    c4.metric("Hari Terpakai", df["Tanggal"].nunique())

    if st.session_state["jumlah_grup"] > st.session_state["kapasitas"]:

        st.warning(
            f"Jumlah kelompok ({st.session_state['jumlah_grup']}) melebihi kapasitas "
            f"({st.session_state['kapasitas']}). Sebagian sesi terisi lebih dari satu "
            "kelompok. Perlebar rentang tanggal, tambah shift, atau tambah kelompok paralel."
        )

    tab1, tab2, tab3 = st.tabs(["📋 Preview Jadwal", "🗓️ Kalender Ringkas", "📊 Rekap"])

    with tab1:

        pilihan_sesi = st.session_state["shifts"]

        filter_sesi = st.multiselect(
            "Filter berdasarkan sesi", options=pilihan_sesi, default=pilihan_sesi
        )

        df_filter = df[df["Jadwal"].isin(filter_sesi)]

        st.dataframe(
            df_filter.drop(columns=["Baris Excel"]),
            use_container_width=True,
            hide_index=True,
        )

    with tab2:

        kalender = (
            df.groupby(["Tanggal", "Hari", "Jadwal"])["Nama Operator"]
            .count()
            .unstack("Jadwal", fill_value=0)
            .reset_index()
        )

        kalender["_urut"] = pd.to_datetime(kalender["Tanggal"], format="%d.%m.%Y")
        kalender = kalender.sort_values("_urut").drop(columns="_urut")

        st.caption("Angka menunjukkan jumlah operator pada tiap sesi.")

        st.dataframe(kalender, use_container_width=True, hide_index=True)

    with tab3:

        rekap = (
            df.groupby(["Tanggal", "Jadwal"])
            .agg(
                Jumlah_Operator=("Nama Operator", "count"),
                Jumlah_SPBU=("Nomor SPBU", "nunique"),
                Jumlah_Kelompok=("Kelompok", "nunique"),
            )
            .reset_index()
        )

        st.dataframe(rekap, use_container_width=True, hide_index=True)

    judul_langkah("5", "Download File Excel")

    st.download_button(
        label="⬇️ Download Excel - Hasil Penjadwalan",
        data=st.session_state["output"],
        file_name="Jadwal_Operator_SPBU_Hasil.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.caption("Kolom jadwal dan tanggal sudah menggunakan Merge Cells per kelompok.")

else:

    st.info(
        "Atur jadwal di sidebar, upload file Excel, lalu klik Generate untuk melihat hasilnya."
    )
