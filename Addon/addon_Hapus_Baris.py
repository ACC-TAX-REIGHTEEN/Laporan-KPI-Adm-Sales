import openpyxl
import os
import sys
import glob

print("--> Program Hapus Baris Excel")

file_list = glob.glob("*.xlsx")

if not file_list:
    print("--> Tidak ditemukan file Excel (.xlsx) di folder ini")
    sys.exit()

print("--> Daftar file yang ditemukan:")
for i, file_name in enumerate(file_list):
    print(f"--> {i + 1}. {file_name}")

try:
    pilihan_file = int(input("--> Pilih nomor file: "))
    if 1 <= pilihan_file <= len(file_list):
        file_path = file_list[pilihan_file - 1]
    else:
        print("--> Nomor pilihan tidak valid")
        sys.exit()
except ValueError:
    print("--> Input harus berupa angka")
    sys.exit()

print(f"--> Membaca file {file_path}...")

try:
    wb = openpyxl.load_workbook(file_path)
except Exception as e:
    print(f"--> Gagal membuka file: {e}")
    sys.exit()

sheet_list = wb.sheetnames

print("--> Daftar Sheet yang tersedia:")
for i, sheet in enumerate(sheet_list):
    print(f"--> {i + 1}. {sheet}")

try:
    pilihan_sheet = int(input("--> Pilih nomor sheet: "))
    if 1 <= pilihan_sheet <= len(sheet_list):
        sheet_name = sheet_list[pilihan_sheet - 1]
    else:
        print("--> Nomor pilihan tidak valid")
        sys.exit()
except ValueError:
    print("--> Input harus berupa angka")
    sys.exit()

ws = wb[sheet_name]

try:
    start_row = int(input("--> Hapus baris mulai: "))
    end_row = int(input("--> Sampai baris: "))
except ValueError:
    print("--> Input harus berupa angka")
    sys.exit()

jumlah_hapus = end_row - start_row + 1

if jumlah_hapus > 0:
    print(f"--> Menghapus {jumlah_hapus} baris di sheet '{sheet_name}'...")
    ws.delete_rows(idx=start_row, amount=jumlah_hapus)
    
    print("--> Menyimpan file...")
    wb.save(file_path)
    print("--> Proses selesai")
else:
    print("--> Range baris tidak valid")
