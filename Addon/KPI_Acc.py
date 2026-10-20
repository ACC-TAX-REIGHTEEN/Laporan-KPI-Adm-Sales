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

kamus_bulan_singkat = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "Mei",
    6: "Jun",
    7: "Jul",
    8: "Agu",
    9: "Sep",
    10: "Okt",
    11: "Nov",
    12: "Des",
}

df["Periode_Sort"] = df["Tgl Waktu Modifikasi"].dt.to_period("M")

periode_terurut = sorted(df["Periode_Sort"].dropna().unique())

kategori_terurut = [
    f"{kamus_bulan_singkat[p.month]} {p.year}" for p in periode_terurut
]

df["Bulan_Tahun"] = (
    df["Tgl Waktu Modifikasi"].dt.month.map(kamus_bulan_singkat)
    + " "
    + df["Tgl Waktu Modifikasi"].dt.year.astype(str)
)

df["Bulan_Tahun"] = pd.Categorical(
    df["Bulan_Tahun"], categories=kategori_terurut, ordered=True
)

tabel_pivot = pd.pivot_table(
    df,
    index=["Nama Pengguna Pengguna", "Tipe transaksi"],
    columns="Bulan_Tahun",
    values="Jumlah Data",
    aggfunc="sum",
    fill_value=0,
    observed=True,
)

tabel_pivot.to_excel(nama_file_output)

print(
    f"--> Selesai! Data berhasil diperbarui dan disimpan dengan nama: {nama_file_output}"
)