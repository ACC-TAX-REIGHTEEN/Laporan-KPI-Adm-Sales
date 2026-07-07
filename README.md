# 📋 Laporan KPI Adm Sales

> **Otomasi pengolahan Form Revisi dokumen menjadi Laporan KPI terstruktur per Depo dan PIC — dari data mentah banyak-entry ke pivot report siap presentasi**

Pipeline Python interaktif yang membaca **`Form Revisi.xlsx`** (log harian revisi dokumen oleh staf Adm Sales), membersihkan data, memecah baris multi-dokumen menjadi satu baris per transaksi, mengoreksi nama dan depo via fuzzy matching, lalu menghasilkan **`Laporan KPI Revisi.xlsx`** berisi ringkasan kinerja revisi per Adm Sales per bulan, pivot per PIC, dan data bersih — atau secara opsional hanya output data yang sudah dibersihkan tanpa pivot.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Persiapan: Format `Form Revisi.xlsx`](#-persiapan-format-form-revisixxxlsx)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Konfigurasi `helper.conf`](#-konfigurasi-helperconf)
- [Output & Struktur Laporan](#-output--struktur-laporan)
- [Alat Bantu (Addon)](#-alat-bantu-addon)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

Form Revisi adalah log harian yang diisi oleh staf Administrasi Sales setiap kali mereka memproses revisi dokumen (Surat Jalan, Invoice, Sales Order, dst.). Masalah umum pada form semacam ini:

- Satu baris bisa berisi **banyak nomor dokumen sekaligus** (dipisah koma, titik koma, atau baris baru)
- **Nama staf dan nama depo** ditulis tidak konsisten (disingkat, typo, atau tercampur kode nomor)
- Kolom tanggal berisi **format campuran** atau nilai tidak valid

Pipeline ini menyelesaikan ketiga masalah tersebut secara otomatis, lalu menghasilkan laporan KPI yang menghitung jumlah revisi per jenis dokumen, per staf, per bulan, dan per Depo.

---

## ✨ Fitur Utama

- **Dua mode output interaktif** — Pilih antara laporan lengkap dengan pivot (Mode 1) atau data bersih mentah tanpa pivot (Mode 2) langsung saat runtime.
- **Pemecahan baris otomatis (row splitting)** — Baris dengan banyak nomor dokumen dipecah menjadi satu baris per nomor, dengan deteksi separator: newline, koma, titik koma, ampersand, dan tanda plus.
- **Deteksi jenis dokumen kontekstual** — Jenis dokumen (SJ, INV, SO, dll.) dideteksi per chunk nomor, bukan per baris keseluruhan — memungkinkan satu baris berisi campuran jenis dokumen yang berbeda.
- **Normalisasi sinonim dokumen** — SI → INV, Job Costing → JC, Retur Penjualan → SR — diterapkan secara otomatis sebelum splitting.
- **Fuzzy matching nama & depo** — Menggunakan `fuzzywuzzy` dengan threshold 80% untuk mengoreksi penulisan tidak konsisten terhadap master data di `helper.conf`.
- **Smart date parser & forward-fill** — Mendukung 5 format tanggal berbeda; tanggal kosong/tidak valid diisi dari nilai baris sebelumnya (atau berikutnya).
- **Exact mapping override** — Kode-kode seperti `Ar-251006967` atau `Mgl` dipetakan ke nama resmi via `[MAPPING_NAMES]` dan `[MAPPING_DEPO]` di `helper.conf`.
- **Validasi rentang tanggal** — Baris dengan tanggal di luar `start_date`–`end_date` diganti dengan tanggal baris sebelumnya.
- **Row splitting berbasis keyword per sales tertentu** — Untuk sales di `target_sales`, baris dengan `Alasan Revisi` berisi beberapa kode dokumen dipecah menjadi satu baris per kode.
- **KPI pivot per Depo** — Setiap depo mendapat blok tabel tersendiri dengan breakdown bulanan per Adm Sales dan ringkasan kumulatif, diberi kode warna unik.
- **Validasi awal menyeluruh** — Mengecek keberadaan semua file dependensi dan memastikan `Form Revisi.xlsx` hanya memiliki satu sheet sebelum proses dimulai.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install openpyxl pandas numpy xlsxwriter fuzzywuzzy python-Levenshtein
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `openpyxl` | A1_A, A1, A2_CosmeticsB, A3 | Baca/tulis Excel, salin style sel |
| `pandas` | A2_Cosmetics, A3 | Baca Excel, groupby, pivot, ExcelWriter |
| `numpy` | A3 | `np.where` untuk klasifikasi biner |
| `xlsxwriter` | A3 | Engine penulisan Excel dengan format warna |
| `fuzzywuzzy` | A2_Cosmetics | Fuzzy string matching nama & depo |
| `python-Levenshtein` | A2_Cosmetics | Akselerasi fuzzywuzzy (opsional, disarankan) |
| `re`, `configparser`, `copy`, `datetime`, `glob`, `shutil`, `subprocess`, `os` | Semua | Standard library |

> **Catatan:** `python-Levenshtein` tidak wajib tapi sangat disarankan karena mempercepat proses fuzzy matching secara signifikan. Tanpanya, `fuzzywuzzy` akan menampilkan peringatan performa.

---

## 📁 Struktur Folder & File

```
📦 Laporan-KPI-Adm-Sales/
│
├── 📄 Hasilkan Laporan Revisi.py    ← Orkestrator utama. Jalankan ini
├── 📄 Form Revisi.xlsx              ← [INPUT] Form revisi dokumen Adm Sales
│
└── 📁 Dapur/                        ← Folder pipeline (jangan diubah strukturnya)
    ├── 📄 __init__.py
    ├── 📄 A1_A_HandleDateOrigin.py  ← Step 1a: Konversi timestamp kolom pertama
    ├── 📄 A1_CleaningNSplitting.py  ← Step 1b: Normalisasi & pecah baris
    ├── 📄 A2_Cosmetics.py           ← Step 2a: Fuzzy matching nama & depo
    ├── 📄 A2_CosmeticsB.py          ← Step 2b: Exact mapping & split keyword
    ├── 📄 A3_PivotData.py           ← Step 3: Pivot KPI report (Mode 1 saja)
    ├── 📄 addon_Hapus_Baris.py      ← Alat bantu: hapus baris range
    ├── 📄 addon_extract_excelSheet.py ← Alat bantu: ekstrak sheet ke file
    └── 📄 helper.conf               ← Konfigurasi master data & mapping
```

> **Output** muncul di folder utama (sejajar dengan `Hasilkan Laporan Revisi.py`):
> - Mode 1 → `Laporan KPI Revisi.xlsx`
> - Mode 2 → `Laporan Revisi Mentah.xlsx`

---

## 📝 Persiapan: Format `Form Revisi.xlsx`

File input **harus memiliki tepat satu sheet**. Orkestrator akan menolak file dengan lebih dari satu sheet.

### Kolom yang dibutuhkan

| Kolom | Keterangan |
|---|---|
| Kolom A | Timestamp mentah dari sistem (format `DD/MM/YYYY HH:MM:SS`). Isinya akan dikonversi ke format tanggal bersih oleh Step 1a. |
| `Tgl Revisi` | Tanggal revisi. Dapat diisi oleh konversi Step 1a atau sudah ada sebelumnya. |
| `Revisi` | Jenis dokumen yang direvisi: `SJ`, `INV`, `SO`, `JV`, `AR`, `LPB`, `TFB`, `FB`, `JC`, `SR`. Sinonim `SI`, `Job Costing`, `Retur Penjualan` juga diterima. |
| `Nomor Revisi` | Nomor dokumen. **Boleh berisi banyak nomor sekaligus** dipisah koma, titik koma, newline, `&`, atau `+`. Nomor valid minimal 5 digit. |
| `Nama Adm Sales` | Nama staf Adm Sales. Penulisan tidak konsisten akan dikoreksi via fuzzy matching. |
| `Depo` | Nama depo/cabang. Dikoreksi via fuzzy matching dan exact mapping. |
| `Divisi` | Divisi (akan diubah ke uppercase). |
| `PIC REVISI` | Person in Charge revisi (akan diubah ke title case). |
| `Alasan Revisi` | Alasan revisi. Digunakan untuk keyword-based splitting pada `target_sales`. |
| `Kesalahan Admin/ Bukan` | Nilai: `Kesalahan Admin` atau nilai lainnya. |

### Jenis Dokumen yang Dikenali

| Kode | Nama Dokumen | Kode Sinonim |
|---|---|---|
| `SJ` | Surat Jalan | — |
| `INV` | Invoice | `SI` |
| `SO` | Sales Order | — |
| `JV` | Journal Voucher | — |
| `AR` | Account Receivable | — |
| `LPB` | Laporan Penerimaan Barang | — |
| `TFB` | Transfer Barang | — |
| `FB` | Faktur Barang | — |
| `JC` | Job Costing | `Job Costing` |
| `SR` | Sales Return | `Retur Penjualan` |

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

Letakkan `Form Revisi.xlsx` di folder utama (sejajar dengan `Hasilkan Laporan Revisi.py`). Pastikan file hanya memiliki **satu sheet aktif**.

### Langkah 2 — Sesuaikan `helper.conf`

Buka `Dapur/helper.conf` dan perbarui:
- `list_depo` — daftar nama depo resmi (pisah koma)
- `list_nama_sales` — daftar nama Adm Sales resmi (pisah koma)
- `start_date` / `end_date` — rentang periode laporan
- `target_sales` — nama sales yang mendapat perlakuan keyword-splitting khusus
- `[MAPPING_NAMES]` dan `[MAPPING_DEPO]` — koreksi eksak untuk nilai yang tidak cocok fuzzy

### Langkah 3 — Jalankan

```bash
python "Hasilkan Laporan Revisi.py"
```

### Langkah 4 — Pilih mode output

```
--> Pilihan Proses:
--> 1. Olah data full dengan python
--> 2. Cukup bersihkan data dan split data
--> Masukkan pilihan (1/2):
```

| Pilihan | Proses yang dijalankan | Output |
|---|---|---|
| `1` | A1_A → A1 → A2 → A2B → A3 | `Laporan KPI Revisi.xlsx` (pivot lengkap) |
| `2` | A1_A → A1 → A2 → A2B | `Laporan Revisi Mentah.xlsx` (data bersih, tanpa pivot) |

### Langkah 5 — Ambil hasil

File output muncul di folder utama. Tekan Enter saat diminta untuk keluar.

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Hasilkan Laporan Revisi.py]
   │
   ├─── Validasi awal
   │       Cek Form Revisi.xlsx ada
   │       Cek Form Revisi.xlsx hanya punya 1 sheet → tolak jika lebih
   │       Cek folder Dapur/ ada
   │       Cek semua file dependensi lengkap (9 file)
   │
   ├─── Bersihkan Dapur/ dari *.xlsx sisa run sebelumnya
   ├─── Salin Form Revisi.xlsx → Dapur/
   ├─── Tampilkan menu pilihan (1/2) → tunggu input pengguna
   │
   ├─── [A1_A] HandleDateOrigin
   │       Input: Form Revisi.xlsx (di Dapur/)
   │       ├─ Baca kolom A (baris 2 dst.)
   │       ├─ Parse nilai datetime: jika sudah datetime → langsung pakai
   │       │   jika string → parse format "DD/MM/YYYY HH:MM:SS"
   │       └─ Tulis tanggal bersih ke kolom E (format DD/MM/YYYY)
   │       Output: Form Revisi.xlsx (dimodifikasi in-place)
   │
   ├─── [A1] CleaningNSplitting
   │       Input: Form Revisi.xlsx
   │       ├─ Buang kolom A (kolom pertama dengan timestamp mentah)
   │       ├─ Normalisasi teks: Divisi → UPPERCASE, Nama Adm/PIC/Depo → Title Case
   │       ├─ Petakan sinonim: SI→INV, Job Costing→JC, Retur Penjualan→SR
   │       ├─ Untuk setiap baris:
   │       │   ├─ Deteksi kode dokumen global dari kolom "Revisi"
   │       │   ├─ Pecah kolom "Nomor Revisi" pada separator: \n , ; & +
   │       │   └─ Per chunk:
   │       │       ├─ Ekstrak semua angka ≥ 5 digit
   │       │       ├─ Deteksi kode lokal dari teks chunk (prioritas atas global)
   │       │       └─ Buat satu baris baru per kombinasi nomor × kode dokumen
   │       └─ Salin style sel (font, fill, border, alignment) ke output
   │       Output: Form Revisi_temp.xlsx
   │
   ├─── [A2] Cosmetics (Fuzzy Matching)
   │       Input: Form Revisi_temp.xlsx
   │       ├─ Baca master data dari helper.conf [MASTER_DATA]
   │       ├─ Smart date parser untuk kolom Tgl Revisi (5 format)
   │       │   Forward-fill: tanggal kosong → diisi dari baris sebelumnya
   │       ├─ Fuzzy match kolom Depo terhadap list_depo (threshold 80%)
   │       └─ Fuzzy match kolom Nama Adm Sales terhadap list_nama_sales (threshold 80%)
   │       Output: Form Revisi_cosm_temp.xlsx
   │
   ├─── [A2B] CosmeticsB (Exact Mapping + Keyword Split)
   │       Input: Form Revisi_cosm_temp.xlsx
   │       ├─ Baca helper.conf: [MAPPING_NAMES], [MAPPING_DEPO], [SETTINGS]
   │       ├─ Terapkan exact mapping Nama Adm Sales dari [MAPPING_NAMES]
   │       ├─ Terapkan exact mapping Depo dari [MAPPING_DEPO]
   │       ├─ Validasi Tgl Revisi: jika di luar start_date–end_date
   │       │   → salin nilai tanggal dari baris sebelumnya
   │       └─ Keyword splitting (hanya untuk target_sales):
   │           Jika Alasan Revisi mengandung beberapa kode dokumen (SJ, INV, dll.)
   │           → sisipkan baris baru dengan kode berbeda, salin semua kolom lain
   │       Output: Form Revisi_cosm_2_temp.xlsx
   │
   ├─── [A3] PivotData ← Hanya Mode 1
   │       Input: Form Revisi_cosm_2_temp.xlsx
   │       ├─ Buat kolom biner: AR=1, FB=1, INV=1, dst. (sesuai nilai kolom Revisi)
   │       ├─ Buat kolom Kesalahan_Admin dan Bukan_Kesalahan_Admin
   │       ├─ Generate sheet "Laporan Depo":
   │       │   ├─ Loop per Depo (warna unik per depo dari palet 8 warna)
   │       │   ├─ Tabel: Nama Adm Sales × Periode (bulan) × count per doc type
   │       │   ├─ Baris "TOTAL BULANAN" dan blok "RINGKASAN KUMULATIF PER ADM"
   │       │   └─ Semua blok tersusun vertikal dalam satu sheet
   │       ├─ Generate sheet "PIC":
   │       │   └─ Pivot: PIC REVISI × Periode (bulan) + GRAND TOTAL
   │       └─ Simpan sheet "Data Asli" (data setelah A2B)
   │       Output: Laporan KPI Revisi.xlsx
   │
   ├─── Salin output → folder utama
   └─── Bersihkan Dapur/ dari semua *.xlsx sementara
```

---

## ⚙️ Konfigurasi `helper.conf`

```ini
[SETTINGS]
target_sales = Citra
keywords = SJ, INV, SO, JV, AR, LPB, TFB, FB, JC, SR
date_format = %d/%m/%Y
start_date = 01/01/2025
end_date = 31/01/2026

[MASTER_DATA]
list_depo = Solo, Jogja, Magelang, Kendari, Makassar, Palu, Pati, Purwokerto, Tegal, Semarang
list_nama_sales = Airen, Budi, Citra, Danu

[MAPPING_NAMES]
Ar-251103500 = Novi
Solo = Siska

[MAPPING_DEPO]
Ar-251006967 = Semarang
Mgl = Magelang
Smg = Semarang
Pusat = Semarang
Kdi = Kendari
MKS = Makassar
```

### Penjelasan tiap seksi

**`[SETTINGS]`**

| Key | Keterangan |
|---|---|
| `target_sales` | Nama sales yang mendapat perlakuan keyword-split di `A2_CosmeticsB.py`. Pisah koma jika lebih dari satu. |
| `keywords` | Daftar kode dokumen yang dikenali saat keyword-splitting kolom `Alasan Revisi`. |
| `date_format` | Format tanggal untuk parsing `start_date` dan `end_date`. Default: `%d/%m/%Y`. |
| `start_date` | Batas awal periode laporan. Baris dengan tanggal sebelum ini dikoreksi. |
| `end_date` | Batas akhir periode laporan. Baris dengan tanggal sesudah ini dikoreksi. |

**`[MASTER_DATA]`**

| Key | Keterangan |
|---|---|
| `list_depo` | Daftar nama depo resmi (pisah koma). Digunakan sebagai target fuzzy matching kolom `Depo`. |
| `list_nama_sales` | Daftar nama Adm Sales resmi (pisah koma). Digunakan sebagai target fuzzy matching kolom `Nama Adm Sales`. |

**`[MAPPING_NAMES]`**

Pemetaan **eksak** (exact match, case-sensitive terhadap nilai setelah fuzzy selesai) untuk nilai nama yang tidak akan terkoreksi oleh fuzzy matching. Format: `nilai_di_data = nama_resmi`.

```ini
Ar-251103500 = Novi       ; kode nomor AR → nama nyata
Solo = Siska              ; nama kota yang salah konteks → nama sales
```

**`[MAPPING_DEPO]`**

Pemetaan eksak untuk nilai depo. Format: `nilai_di_data = nama_depo_resmi`.

```ini
Mgl = Magelang
Smg = Semarang
Kdi = Kendari
MKS = Makassar
```

> **Urutan koreksi:** Fuzzy matching (A2) dijalankan **sebelum** exact mapping (A2B). Artinya, jika nilai sudah terkoreksi oleh fuzzy, exact mapping tidak akan menimpa lagi kecuali ada entri yang cocok secara eksak.

---

## 📤 Output & Struktur Laporan

### Mode 1 — `Laporan KPI Revisi.xlsx`

Berisi 3 sheet:

#### Sheet 1 — `Laporan Depo`

Satu sheet vertikal berisi blok untuk **setiap depo** secara berurutan. Setiap blok depo terdiri dari:

**Blok A: Tabel detail bulanan**

Header baris: `Nama Adm Sales | Periode | Kesalahan_Admin | Bukan_Kesalahan_Admin | AR | FB | INV | JC | JV | LPB | SJ | SO | SR | TFB`

- Satu baris per kombinasi Nama Adm Sales × Periode (mis. `Jan 2025`)
- Angka di setiap kolom dokumen = jumlah revisi dokumen tersebut pada periode itu
- Baris terakhir: `TOTAL BULANAN` (grand total seluruh Adm Sales seluruh periode di depo ini)

**Blok B: Ringkasan kumulatif per Adm**

Header: `Nama Adm Sales | Kesalahan_Admin | Bukan_Kesalahan_Admin | TOTAL_REVISI`

- Satu baris per Adm Sales (akumulasi seluruh periode)
- Baris terakhir: `TOTAL KUMULATIF`

Setiap depo mendapat warna judul berbeda (biru, merah, hijau, ungu, cyan, oranye, dst.) secara bergantian dari palet 8 warna.

#### Sheet 2 — `PIC`

Tabel pivot cross-tab:

| PIC REVISI | Jan 2025 | Feb 2025 | ... | TOTAL |
|---|---|---|---|---|
| Nama PIC 1 | 12 | 8 | ... | 43 |
| ... | | | | |
| GRAND TOTAL | | | | |

- Baris = setiap PIC REVISI unik
- Kolom = setiap bulan yang ada di data (diurutkan kronologis)
- Header biru gelap (`#366092`), baris GRAND TOTAL diberi fill biru muda
- Auto-fit lebar kolom

#### Sheet 3 — `Data Asli`

Data lengkap setelah seluruh proses pembersihan (output A2B), tanpa transformasi pivot. Berguna untuk verifikasi manual atau analisis lanjutan.

---

### Mode 2 — `Laporan Revisi Mentah.xlsx`

Data yang sudah melalui 4 step awal (A1_A → A1 → A2 → A2B): baris sudah dipecah satu per dokumen, nama dan depo sudah dikoreksi, tanggal sudah divalidasi — namun **belum di-pivot**. Berguna jika:
- Ingin melakukan filter atau analisis mandiri di Excel/spreadsheet lain
- Ingin mengupload ke Google Sheets atau database
- Ingin memverifikasi hasil pembersihan sebelum membuat pivot

---

## 🛠️ Alat Bantu (Addon)

Dua skrip di folder `Dapur/` yang **tidak dijalankan oleh pipeline** — digunakan secara manual sesuai kebutuhan.

### `addon_Hapus_Baris.py`

Alat interaktif untuk menghapus rentang baris dari file Excel manapun di folder yang sama.

```bash
cd Dapur
python addon_Hapus_Baris.py
# Pilih file → pilih sheet → masukkan baris mulai & sampai
```

Berguna untuk membersihkan baris header berulang, baris total yang tidak diinginkan, atau data yang perlu dikecualikan secara manual sebelum diproses ulang.

### `addon_extract_excelSheet.py`

Alat interaktif untuk mengekstrak satu atau semua sheet dari file Excel menjadi file-file terpisah, dengan format (style, lebar kolom, dll.) tetap terjaga.

```bash
cd Dapur
python addon_extract_excelSheet.py
# Pilih file → pilih sheet (atau 0 untuk semua)
# Output: [NamaSheet].xlsx per sheet yang dipilih
```

Berguna jika `Laporan KPI Revisi.xlsx` perlu dibagikan per depo atau per sheet secara terpisah.

---

## 🛠️ Troubleshooting

### ❌ `File "Form Revisi.xlsx" tidak ditemukan`
Pastikan file ada di folder utama (sejajar dengan `Hasilkan Laporan Revisi.py`), dengan nama **persis** `Form Revisi.xlsx` (termasuk spasi, kapital, dan ekstensi `.xlsx`).

### ❌ `File tersebut memiliki banyak sheet, hapus dan sisakan yang akan di olah`
`Form Revisi.xlsx` memiliki lebih dari satu sheet. Buka file di Excel, hapus semua sheet kecuali yang berisi data yang akan diproses, lalu simpan ulang.

### ❌ `Gagal: File berikut tidak ditemukan di folder Dapur`
Salah satu atau lebih skrip / file konfigurasi di `Dapur/` terhapus atau tidak ada. Pastikan seluruh file yang tercantum di bagian Struktur Folder tersedia.

### ❌ `Gagal: File konfigurasi helper.conf tidak ditemukan` (dari A2_Cosmetics)
`helper.conf` harus berada di `Dapur/`. Periksa path dan pastikan file tidak dipindahkan.

### ❌ Semua nilai Depo / Nama Adm Sales tetap tidak berubah setelah proses
Fuzzy matching tidak menemukan kandidat dengan skor ≥ 80%. Kemungkinan penyebab:
- Nilai di data sangat berbeda dari `list_depo` / `list_nama_sales` di `helper.conf`
- Master data belum diisi dengan nilai yang representatif
Solusi: tambahkan aturan eksak di `[MAPPING_DEPO]` / `[MAPPING_NAMES]` untuk kasus yang tidak bisa di-fuzzy.

### ❌ Kolom pivot di `Laporan Depo` bernilai 0 semua untuk satu kode dokumen
Berarti tidak ada baris data dengan nilai tersebut di kolom `Revisi`. Pastikan proses A1 berhasil mengisi kolom `Revisi` dengan kode yang valid (bukan sinonim yang belum dikonversi).

### ❌ `ModuleNotFoundError: No module named 'fuzzywuzzy'`
```bash
pip install fuzzywuzzy python-Levenshtein
```

### ❌ Sheet "PIC" kosong atau hanya ada header
Kolom `PIC REVISI` tidak ditemukan atau semua nilainya kosong. Pastikan nama kolom di `Form Revisi.xlsx` persis `PIC REVISI` (kapital semua, ada spasi).

### ❌ Baris hasil splitting sangat banyak / tidak masuk akal
Terjadi jika kolom `Nomor Revisi` berisi teks panjang dengan banyak angka 5+ digit yang bukan nomor dokumen (misalnya nomor telepon atau kode lain). Periksa data sumber dan pastikan kolom `Nomor Revisi` hanya berisi nomor dokumen.

---

## 📌 Catatan Penting

- **Form Revisi.xlsx wajib satu sheet** — Ini adalah validasi wajib di awal. Jika file punya lebih dari satu sheet (misal ada sheet ringkasan), proses akan dihentikan.
- **File input tidak dimodifikasi** — Orkestrator menyalin (bukan memindahkan) `Form Revisi.xlsx` ke `Dapur/`. File asli di folder utama tetap utuh.
- **Output lama akan ditimpa** — Jika `Laporan KPI Revisi.xlsx` atau `Laporan Revisi Mentah.xlsx` sudah ada dari run sebelumnya, file lama langsung diganti. Simpan salinan jika diperlukan.
- **Jangan ubah struktur `Dapur/`** — Semua skrip menggunakan path relatif dan saling bergantung pada nama file temp yang sudah ditentukan (`Form Revisi_temp.xlsx`, `Form Revisi_cosm_temp.xlsx`, dst.).
- **Urutan mapping: fuzzy dulu, exact kemudian** — A2_Cosmetics menjalankan fuzzy matching, lalu A2_CosmeticsB menjalankan exact mapping. Nilai yang sudah benar setelah fuzzy tidak akan diubah lagi oleh exact mapping (kecuali ada entri yang persis sama di `[MAPPING_NAMES]`/`[MAPPING_DEPO]`).
- **Keyword splitting A2B hanya untuk `target_sales`** — Baris dari Adm Sales yang namanya tidak ada di `target_sales` di `helper.conf` tidak akan dipecah berdasarkan keyword, meski `Alasan Revisi`-nya mengandung banyak kode dokumen.
- **Warna depo tidak persisten** — Urutan warna di `Laporan Depo` bergantung pada urutan depo unik yang muncul di data. Depo yang sama bisa mendapat warna berbeda jika dijalankan dengan data yang berbeda.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perpajakan dan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
