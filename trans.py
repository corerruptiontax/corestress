import pandas as pd
import os
from datetime import datetime
from colorama import Fore, Style  # Tambah impor ini
from src.trans.file_reader import read_raw_file
from src.trans.data_processor import process_data
from src.trans.excel_formatter import save_to_excel
from src.trans.logger import log_important, log_info, log_success, log_error, log_summary
from src.trans.format_normalizer import normalize_format

def main():
    try:
        log_info("Memulai proses transformasi data...")
        raw_file_name = input(f"{Fore.CYAN}Masukkan nama file raw (contoh: 'xxxx - Raw'): {Style.RESET_ALL}").strip()
        sheet_name = input(f"{Fore.CYAN}Masukkan nama sheet di file raw (kosongkan untuk sheet pertama): {Style.RESET_ALL}").strip() or None
        process_file_name = input(f"{Fore.CYAN}Masukkan nama file proses (contoh: 'xxxx - Process'): {Style.RESET_ALL}").strip()

        columns = ['No. Pelanggan', 'Nama Pelanggan', 'No. Faktur', 'Tgl. Faktur', 'Nama Barang', 
                   'Harga DPP', 'Qty', 'Total DPP', 'PPN', 'Baris']

        # Baca file raw dengan nama sheet yang ditentukan
        log_info(f"Membaca file raw: {raw_file_name}")
        df = read_raw_file(raw_file_name, sheet_name)

        # Normalisasi format data dan dapatkan informasi header asli
        log_info("Menormalisasi format data...")
        df, is_dpp_header = normalize_format(df)

        # Proses data dan ambil hasil, dengan informasi header asli
        log_info("Memproses data...")
        result_data, deleted_zero, deleted_minus, deleted_rows, total_dpp, total_ppn = process_data(df, is_dpp_header)

        # Simpan hasil utama ke file Excel
        log_info(f"Menyimpan hasil ke: {process_file_name}")
        save_to_excel(result_data, process_file_name, columns)
        log_success(f"File hasil berhasil disimpan sebagai: {process_file_name}")

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
            log_success(f"Baris yang dihapus disimpan ke: {deleted_file_path}")

        # Tampilkan laporan ringkasan
        log_summary(f"Laporan Ringkasan: Total DPP = {total_dpp:.2f}, Total PPN = {total_ppn:.2f}")
        log_summary(f"Total baris yang dihapus karena nilai nol: {deleted_zero}")
        log_summary(f"Total baris yang dihapus karena nilai minus (retur): {deleted_minus}")

        # Tampilkan timestamp
        timestamp = datetime.now().strftime("%I:%M %p WIB, %A, %d %B %Y")
        log_success(f"Transformasi selesai pada {timestamp}!")

    except Exception as e:
        log_error(f"Terjadi kesalahan: {str(e)}")

if __name__ == "__main__":
    main()