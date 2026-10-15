import os
import pandas as pd

nama_file_asli = "combine.xlsx"

nama_dasar, ekstensi = os.path.splitext(nama_file_asli)
nama_file_output = f"{nama_dasar}_pivot{ekstensi}"

print(f"--> Membaca file: {nama_file_asli}...")

df = pd.read_excel(nama_file_asli)

df["Tgl Waktu Modifikasi"] = pd.to_datetime(
    df["Tgl Waktu Modifikasi"], dayfirst=True
)

kamus_bulan = {
    1: "Januari",
    2: "Februari",
    3: "Maret",
    4: "April",
    5: "Mei",
    6: "Juni",
    7: "Juli",
    8: "Agustus",
    9: "September",
    10: "Oktober",
    11: "November",
    12: "Desember",
}

df["Bulan"] = df["Tgl Waktu Modifikasi"].dt.month.map(kamus_bulan)

urutan_bulan = list(kamus_bulan.values())
df["Bulan"] = pd.Categorical(df["Bulan"], categories=urutan_bulan, ordered=True)

tabel_pivot = pd.pivot_table(
    df,
    index=["Nama Pengguna Pengguna", "Tipe transaksi"],
    columns="Bulan",
    values="Jumlah Data",
    aggfunc="sum",
    fill_value=0,
    observed=True,
)

tabel_pivot.to_excel(nama_file_output)

print(
    f"--> Selesai! Data berhasil diperbarui dan disimpan dengan nama: {nama_file_output}"
)
