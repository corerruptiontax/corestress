# cust.py
from openpyxl import load_workbook
import pandas as pd
from colorama import Fore, Style
from tqdm import tqdm
import os
from utils import convert_date

def process_customer(template_file, source_file, use_ref, id_tku):
    """
    Memproses data customer dari file sumber ke template Excel
    
    Parameter:
    template_file (str): Path file template Excel
    source_file (str): Path file sumber data customer
    use_ref (bool): Flag untuk menggunakan kolom referensi
    id_tku (str): ID TKU sesuai lokasi
    """
    print(Fore.CYAN + "\n=== PROSES DATA CUSTOMER ===" + Style.RESET_ALL)
    
    try:
        # Validasi file
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"File sumber '{source_file}' tidak ditemukan!")
            
        if not source_file.endswith('.xlsx'):
            source_file += '.xlsx'

        # Baca data
        print(Fore.BLUE + "🔍 Membaca file sumber..." + Style.RESET_ALL)
        df = pd.read_excel(source_file, sheet_name=0)
        
        # Validasi kolom
        req_cols = ['Nama Pelanggan', 'Tgl. Faktur', 'No. Pelanggan']
        missing = [col for col in req_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Kolom wajib tidak ditemukan: {missing}")

        # Buka template
        print(Fore.BLUE + "📂 Membuka template Excel..." + Style.RESET_ALL)
        wb = load_workbook(template_file)
        sheet = wb['Faktur']

        # Proses data
        print(Fore.BLUE + "🔄 Memproses baris data..." + Style.RESET_ALL)
        current_row = 4
        for _, row in tqdm(df.iterrows(), desc="Progress", unit="row"):
            if pd.isna(row['Nama Pelanggan']) or pd.isna(row['Tgl. Faktur']):
                continue

            # Konversi tanggal
            tgl = convert_date(row['Tgl. Faktur'])
            
            # Tulis data
            sheet[f'A{current_row}'] = current_row - 3
            sheet[f'B{current_row}'] = tgl
            sheet[f'C{current_row}'] = 'Normal'
            sheet[f'D{current_row}'] = '04'
            sheet[f'L{current_row}'] = 'IDN'
            sheet[f'R{current_row}'] = row['No. Pelanggan']
            sheet[f'N{current_row}'] = row['Nama Pelanggan']
            sheet[f'I{current_row}'] = id_tku
            
            if use_ref:
                sheet[f'G{current_row}'] = row.get('No. Faktur', '')
                
            current_row += 1

        # Simpan
        print(Fore.BLUE + "💾 Menyimpan perubahan..." + Style.RESET_ALL)
        wb.save(template_file)
        print(Fore.GREEN + f"✅ Sukses! Total data diproses: {current_row-4}" + Style.RESET_ALL)

    except Exception as e:
        error_msg = f"""
        ❌ ERROR PROSES CUSTOMER:
        File: {os.path.basename(source_file)}
        Baris Error: {current_row if 'current_row' in locals() else 'N/A'}
        Detil Error: {str(e)}
        """
        print(Fore.RED + error_msg + Style.RESET_ALL)
