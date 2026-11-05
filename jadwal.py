import streamlit as st
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from datetime import date, timedelta
from collections import OrderedDict
from io import BytesIO

# =========================================================
# 1. KONFIGURASI STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Penjadwalan Operator SPBU", page_icon="📅", layout="wide"
)

st.title("📅 Penjadwalan Pelatihan Operator SPBU")

st.markdown("""
    Aplikasi untuk membuat jadwal pelatihan operator SPBU
    secara otomatis berdasarkan nomor SPBU, dengan
    penggabungan dua operator dalam satu sesi pelatihan.
    """)


# =========================================================
# 2. KONFIGURASI JADWAL
# =========================================================

SESI = ["Sesi Pagi (08.00-11.00)", "Sesi Siang (13.00-16.00)"]

KOLOM_SPBU = 3  # C
KOLOM_OPERATOR = 4  # D
KOLOM_JADWAL = 8  # H
KOLOM_TANGGAL = 9  # I

BARIS_AWAL = 2


# =========================================================
# 3. FUNGSI TANGGAL ROTASI
# =========================================================


def tanggal_rotasi(index, mulai, maksimum):

    total_hari = (maksimum - mulai).days + 1

    return mulai + timedelta(days=index % total_hari)


# =========================================================
# 4. FUNGSI MEMBUAT KELOMPOK OPERATOR
# =========================================================


def buat_kelompok_operator(ws):

    kelompok = OrderedDict()

    for row in range(BARIS_AWAL, ws.max_row + 1):

        nomor_spbu = ws.cell(row, KOLOM_SPBU).value

        nama_operator = ws.cell(row, KOLOM_OPERATOR).value

        if nomor_spbu is None or nama_operator is None:
            continue

        if not str(nama_operator).strip():
            continue

        nomor_spbu = str(nomor_spbu).strip()

        if nomor_spbu not in kelompok:
            kelompok[nomor_spbu] = []

        kelompok[nomor_spbu].append(
            {"row": row, "spbu": nomor_spbu, "operator": str(nama_operator).strip()}
        )

    return kelompok


# =========================================================
# 5. FUNGSI MEMASANGKAN OPERATOR
# =========================================================


def buat_pasangan(kelompok):

    pasangan = []
    operator_sisa = None

    for nomor_spbu, operators in kelompok.items():

        daftar = []

        # Operator sisa dari SPBU sebelumnya
        if operator_sisa is not None:
            daftar.append(operator_sisa)
            operator_sisa = None

        daftar.extend(operators)

        i = 0

        while i + 1 < len(daftar):

            pasangan.append([daftar[i], daftar[i + 1]])

            i += 2

        # Jika ganjil, simpan operator terakhir
        if i < len(daftar):
            operator_sisa = daftar[i]

    if operator_sisa is not None:

        raise ValueError(
            f"Operator {operator_sisa['operator']} "
            "belum memiliki pasangan. "
            "Jumlah keseluruhan operator harus genap."
        )

    return pasangan


# =========================================================
# 6. FUNGSI MEMPROSES JADWAL DAN MERGE CELLS
# =========================================================


def proses_penjadwalan(file_bytes, nama_sheet, tanggal_mulai, tanggal_maksimum):

    wb = load_workbook(BytesIO(file_bytes))

    if nama_sheet not in wb.sheetnames:
        raise ValueError("Sheet tidak ditemukan.")

    ws = wb[nama_sheet]

    # -----------------------------------------------------
    # MEMBACA DATA
    # -----------------------------------------------------

    kelompok = buat_kelompok_operator(ws)

    if not kelompok:
        raise ValueError("Tidak ditemukan data operator pada kolom C dan D.")

    # -----------------------------------------------------
    # MEMBUAT PASANGAN
    # -----------------------------------------------------

    pasangan = buat_pasangan(kelompok)

    if not pasangan:
        raise ValueError("Tidak ada pasangan operator.")

    # -----------------------------------------------------
    # VALIDASI BARIS PASANGAN
    # -----------------------------------------------------

    # Merge vertikal membutuhkan dua baris berurutan.
    # Validasi dilakukan sebelum mengubah workbook.

    for pair in pasangan:

        row1 = pair[0]["row"]
        row2 = pair[1]["row"]

        if abs(row2 - row1) != 1:

            raise ValueError(
                f"Pasangan {pair[0]['operator']} dan "
                f"{pair[1]['operator']} berada pada "
                f"baris {row1} dan {row2}, tidak berurutan. "
                "Susun ulang data operator agar setiap "
                "pasangan berada pada baris bersebelahan."
            )

    # -----------------------------------------------------
    # MELEPAS MERGE LAMA PADA KOLOM H DAN I
    # -----------------------------------------------------

    merged_ranges = list(ws.merged_cells.ranges)

    for merged in merged_ranges:

        if (
            merged.min_col <= KOLOM_TANGGAL
            and merged.max_col >= KOLOM_JADWAL
            and merged.max_row >= BARIS_AWAL
        ):

            ws.unmerge_cells(str(merged))

    # -----------------------------------------------------
    # MEMBERSIHKAN NILAI LAMA KOLOM H DAN I
    # -----------------------------------------------------

    for row in range(BARIS_AWAL, ws.max_row + 1):

        ws.cell(row, KOLOM_JADWAL).value = None
        ws.cell(row, KOLOM_TANGGAL).value = None

    # -----------------------------------------------------
    # MEMBUAT JADWAL
    # -----------------------------------------------------

    hasil = []

    for index, pair in enumerate(pasangan):

        tanggal = tanggal_rotasi(index, tanggal_mulai, tanggal_maksimum)

        sesi = SESI[index % 2]

        row1 = pair[0]["row"]
        row2 = pair[1]["row"]

        row_awal = min(row1, row2)
        row_akhir = max(row1, row2)

        # ---------------------------------------------
        # KOLOM H: JADWAL
        # ---------------------------------------------

        ws.cell(row_awal, KOLOM_JADWAL).value = sesi

        # ---------------------------------------------
        # KOLOM I: TANGGAL
        # ---------------------------------------------

        ws.cell(row_awal, KOLOM_TANGGAL).value = tanggal

        ws.cell(row_awal, KOLOM_TANGGAL).number_format = "dd.mm.yyyy"

        # ---------------------------------------------
        # MERGE KOLOM H DAN I
        # ---------------------------------------------

        ws.merge_cells(
            start_row=row_awal,
            start_column=KOLOM_JADWAL,
            end_row=row_akhir,
            end_column=KOLOM_JADWAL,
        )

        ws.merge_cells(
            start_row=row_awal,
            start_column=KOLOM_TANGGAL,
            end_row=row_akhir,
            end_column=KOLOM_TANGGAL,
        )

        # ---------------------------------------------
        # ALIGNMENT
        # ---------------------------------------------

        ws.cell(row_awal, KOLOM_JADWAL).alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

        ws.cell(row_awal, KOLOM_TANGGAL).alignment = Alignment(
            horizontal="center", vertical="center"
        )

        # ---------------------------------------------
        # MEMBUAT DATA PREVIEW
        # ---------------------------------------------

        for operator in pair:

            hasil.append(
                {
                    "Nomor SPBU": operator["spbu"],
                    "Nama Operator": operator["operator"],
                    "Jadwal": sesi,
                    "Tanggal": tanggal.strftime("%d.%m.%Y"),
                    "Baris Excel": operator["row"],
                }
            )

    # -----------------------------------------------------
    # MENYIMPAN FILE KE MEMORY
    # -----------------------------------------------------

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return (output.getvalue(), hasil, len(kelompok), len(pasangan))


# =========================================================
# 7. SIDEBAR PENGATURAN
# =========================================================

st.sidebar.header("⚙️ Pengaturan Jadwal")

tanggal_mulai = st.sidebar.date_input("Tanggal Mulai", value=date(2026, 11, 9))

tanggal_maksimum = st.sidebar.date_input("Tanggal Maksimum", value=date(2026, 11, 27))

st.sidebar.divider()

st.sidebar.markdown("**Ketentuan Sesi**")

st.sidebar.info("🌅 Sesi Pagi\n\n" "08.00-11.00\n\n" "🌤️ Sesi Siang\n\n" "13.00-16.00")

st.sidebar.caption("Setiap pasangan terdiri dari dua operator.")


# =========================================================
# 8. UPLOAD EXCEL
# =========================================================

st.header("1. Upload Data Operator")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])


if uploaded_file is not None:

    file_bytes = uploaded_file.getvalue()

    try:

        wb_preview = load_workbook(BytesIO(file_bytes), read_only=True)

        daftar_sheet = wb_preview.sheetnames

        wb_preview.close()

        st.success(f"File berhasil dibaca: {uploaded_file.name}")

        # -------------------------------------------------
        # PILIH SHEET
        # -------------------------------------------------

        st.header("2. Pilih Sheet")

        index_default = (
            daftar_sheet.index("dataoperator(2)")
            if "dataoperator(2)" in daftar_sheet
            else 0
        )

        nama_sheet = st.selectbox(
            "Pilih sheet yang akan diproses", daftar_sheet, index=index_default
        )

        # -------------------------------------------------
        # INFORMASI POLA
        # -------------------------------------------------

        with st.expander("Lihat aturan penjadwalan"):

            st.markdown("""
                1. Operator dikelompokkan berdasarkan nomor SPBU.
                2. Dua operator mendapatkan satu jadwal yang sama.
                3. Operator sisa dibawa ke SPBU berikutnya.
                4. Sesi pagi dan siang bergantian.
                5. Tanggal berputar sesuai rentang yang dipilih.
                6. Kolom H dan I di-merge untuk setiap pasangan.
                """)

        # -------------------------------------------------
        # TOMBOL GENERATE
        # -------------------------------------------------

        st.divider()

        if st.button(
            "🚀 GENERATE JADWAL OTOMATIS", type="primary", use_container_width=True
        ):

            if tanggal_maksimum < tanggal_mulai:

                st.error("Tanggal maksimum harus setelah tanggal mulai.")

            else:

                try:

                    with st.spinner("Sedang membuat jadwal operator..."):

                        output, hasil, jumlah_spbu, jumlah_pasangan = (
                            proses_penjadwalan(
                                file_bytes, nama_sheet, tanggal_mulai, tanggal_maksimum
                            )
                        )

                    # Simpan hasil pada session state
                    st.session_state["output"] = output
                    st.session_state["hasil"] = hasil
                    st.session_state["jumlah_spbu"] = jumlah_spbu
                    st.session_state["jumlah_pasangan"] = jumlah_pasangan

                    st.success("Jadwal berhasil dibuat!")

                except Exception as e:

                    st.error(f"Gagal membuat jadwal: {e}")

    except Exception as e:

        st.error(f"Gagal membaca file: {e}")


# =========================================================
# 9. MENAMPILKAN HASIL
# =========================================================

if "hasil" in st.session_state:

    st.divider()

    st.header("3. Hasil Penjadwalan")

    col1, col2, col3 = st.columns(3)

    col1.metric("Jumlah SPBU", st.session_state["jumlah_spbu"])

    col2.metric("Jumlah Pasangan", st.session_state["jumlah_pasangan"])

    col3.metric("Jumlah Operator", len(st.session_state["hasil"]))

    df = pd.DataFrame(st.session_state["hasil"])

    # -----------------------------------------------------
    # FILTER SESI
    # -----------------------------------------------------

    st.subheader("Preview Jadwal")

    filter_sesi = st.multiselect("Filter berdasarkan sesi", options=SESI, default=SESI)

    df_filter = df[df["Jadwal"].isin(filter_sesi)]

    st.dataframe(
        df_filter.drop(columns=["Baris Excel"]),
        use_container_width=True,
        hide_index=True,
    )

    # -----------------------------------------------------
    # REKAP TANGGAL
    # -----------------------------------------------------

    st.subheader("Rekap Per Tanggal")

    rekap = (
        df.groupby(["Tanggal", "Jadwal"])
        .agg(
            Jumlah_Operator=("Nama Operator", "count"),
            Jumlah_SPBU=("Nomor SPBU", "nunique"),
        )
        .reset_index()
    )

    st.dataframe(rekap, use_container_width=True, hide_index=True)

    # -----------------------------------------------------
    # DOWNLOAD FILE EXCEL
    # -----------------------------------------------------

    st.divider()

    st.header("4. Download File Excel")

    st.download_button(
        label="⬇️ DOWNLOAD EXCEL - HASIL PENJADWALAN",
        data=st.session_state["output"],
        file_name="Jadwal_Operator_SPBU_Hasil.xlsx",
        mime=("application/vnd.openxmlformats-officedocument." "spreadsheetml.sheet"),
        type="primary",
        use_container_width=True,
    )

    st.success(
        "File Excel telah siap diunduh. " "Kolom H dan I sudah menggunakan Merge Cells."
    )

else:

    st.info(
        "Upload file Excel terlebih dahulu, "
        "kemudian pilih sheet dan generate jadwal."
    )
