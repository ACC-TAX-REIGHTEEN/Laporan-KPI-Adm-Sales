# Laporan-KPI-Adm-Sales

Proyek ini adalah sistem automasi berbasis Python untuk mengolah data mentah laporan revisi (biasanya dari Google Forms/Excel) menjadi Laporan KPI yang rapi dan terstruktur.

Sistem ini melakukan pembersihan data (cleaning), perbaikan typo otomatis (fuzzy logic), pemecahan baris data ganda (splitting), hingga pembuatan Pivot Table secara otomatis.

# Fitur Utama
- Pembersihan Data Otomatis: Memperbaiki format tanggal dan menghapus karakter yang tidak perlu.
- Smart Splitting: Memecah satu baris data yang berisi banyak nomor revisi menjadi baris-baris terpisah agar perhitungan akurat.
- Auto-Correction: Menggunakan Fuzzy Logic untuk memperbaiki kesalahan penulisan nama Sales atau Depo (misal: "Semaranggg" -> "Semarang").
- Format Preservation: Menjaga style (warna/bold) dari file asli saat melakukan copy data.
- Pivot Table Reporting: Menghasilkan laporan akhir berupa Pivot Table per Depo dan per PIC dengan conditional formatting.
- Configurable: Daftar nama Adm Sales, Depo, dan Keyword kesalahan dapat diatur melalui file konfigurasi eksternal (helper.conf).

# Prasyarat (Requirements)
Program ini membutuhkan Python 3.x dan beberapa library pendukung.

1. Instalasi Python
Pastikan Python sudah terinstall di komputer. Cek dengan perintah:
```bash
python --version
```

2. Instalasi Library
Salin dan jalankan perintah berikut di terminal/command prompt untuk menginstall semua modul yang dibutuhkan:
```bash
pip install pandas openpyxl xlsxwriter fuzzywuzzy python-Levenshtein
```

# Cara Penggunaan
1. Siapkan Data Input
- Siapkan file Excel data mentah.
- Ubah nama filenya menjadi "Form Revisi.xlsx" (Lihat template Excel).
- Simpan file tersebut di folder utama (sejajar dengan Hasilkan Laporan Revisi.py).

2. Jalankan Program Buka terminal/CMD di folder project, lalu jalankan perintah:
```bash
python "Hasilkan Laporan Revisi.py"
```
Atau dengan 2× klik pada file tersebut.

3. Ikuti Instruksi di Layar Program akan melakukan pengecekan file. Jika semua aman, kamu akan diberi pilihan:
- Pilih 1: Untuk proses lengkap sampai jadi Laporan KPI (Pivot Table).
- Pilih 2: Hanya untuk membersihkan dan merapikan data mentah saja.

4. Hasil Akhir
- Jika sukses (dengan pilihan 1), maka file Laporan KPI Revisi.xlsx akan muncul di folder utama.
- File ini berisi sheet: Laporan Depo, PIC, dan Data Asli.

# Konfigurasi (helper.conf)
Silakan dapat ubah konfigurasi pada file tersebut, seperti: daftar nama karyawan, depo, atau periode laporan tanpa menyentuh kode Python. Cukup edit file Dapur/helper.conf.
