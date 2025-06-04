import pandas as pd
import os
from datetime import datetime
from src.trans.file_reader import read_raw_file  # Impor dari src/trans/
from src.trans.data_processor import process_data  # Impor dari src/trans/
from src.trans.excel_formatter import save_to_excel  # Impor dari src/trans/
from src.trans.logger import log_important, log_info  # Impor dari src/trans/
from src.trans.format_normalizer import normalize_format  # Impor dari src/trans/

def main():
    raw_file_name = input("Masukkan nama file raw (contoh: 'xxxx - Raw'): ")
    sheet_name = input("Masukkan nama sheet di file raw (kosongkan untuk sheet pertama): ").strip() or None
    process_file_name = input("Masukkan nama file proses (contoh: 'xxxx - Process'): ")

    columns = ['No. Pelanggan', 'Nama Pelanggan', 'No. Faktur', 'Tgl. Faktur', 'Nama Barang', 
               'Harga DPP', 'Qty', 'Total DPP', 'PPN', 'Baris']

    # Baca file raw dengan nama sheet yang ditentukan
    df = read_raw_file(raw_file_name, sheet_name)

    # Normalisasi format data dan dapatkan informasi header asli
    df, is_dpp_header = normalize_format(df)

    # Proses data dan ambil hasil, dengan informasi header asli
    result_data, deleted_zero, deleted_minus, deleted_rows, total_dpp, total_ppn = process_data(df, is_dpp_header)

    # Simpan hasil utama ke file Excel
    save_to_excel(result_data, process_file_name, columns)

    # Simpan baris yang dihapus ke file terpisah di folder Trans Deleted (jika ada)
    if deleted_rows:
        deleted_columns = columns
        trans_deleted_folder = "Trans Deleted"
        os.makedirs(trans_deleted_folder, exist_ok=True)
        current_date = datetime.now().strftime("%d%m%Y")
        base_name = process_file_name.replace('.xlsx', '')
        deleted_file_name = f"{base_name}_deleted_{current_date}.xlsx"
        deleted_file_path = os.path.join(trans_deleted_folder, deleted_file_name)

        # Rapikan deleted_rows: hapus lompatan ganda baris kosong
        cleaned_deleted_rows = []
        last_was_data = False
        for row in deleted_rows:
            is_empty = row[0] is None  # Baris kosong
            if is_empty:
                if last_was_data:  # Tambahkan baris kosong hanya jika baris sebelumnya adalah data
                    cleaned_deleted_rows.append(row)
                last_was_data = False
            else:
                cleaned_deleted_rows.append(row)
                last_was_data = True

        # Hapus baris kosong di awal atau akhir
        while len(cleaned_deleted_rows) > 0 and cleaned_deleted_rows[0][0] is None:
            cleaned_deleted_rows.pop(0)
        while len(cleaned_deleted_rows) > 0 and cleaned_deleted_rows[-1][0] is None:
            cleaned_deleted_rows.pop(-1)

        # Simpan cleaned_deleted_rows ke file
        with pd.ExcelWriter(deleted_file_path, engine='openpyxl') as writer:
            pd.DataFrame(cleaned_deleted_rows, columns=deleted_columns).to_excel(writer, index=False, sheet_name='Sheet1')
        print(f"Baris yang dihapus telah disimpan ke {deleted_file_path}")

    # Tampilkan total keseluruhan DPP dan PPN
    print(f"\nLaporan Ringkasan: Total DPP = {total_dpp:.2f}, Total PPN = {total_ppn:.2f}")

    # Tampilkan total baris yang dihapus
    log_info(f"Total baris yang dihapus karena nilai nol: {deleted_zero}")
    log_info(f"Total baris yang dihapus karena nilai minus (retur): {deleted_minus}")

    # Tampilkan timestamp
    timestamp = "04:34 PM WIB, Wednesday, 04 June 2025"
    print(f"Transformasi selesai pada {timestamp}! File disimpan sebagai: {process_file_name}")

if __name__ == "__main__":
    main()