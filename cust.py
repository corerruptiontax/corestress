# cust.py
from openpyxl import load_workbook
import pandas as pd
from colorama import Fore, Style

def process_customer(template_file, source_file, use_ref):
    print(Fore.CYAN + "\n=== INPUT DATA CUSTOMER ===" + Style.RESET_ALL)
    try:
        # Memastikan file sumber memiliki ekstensi .xlsx
        if not source_file.endswith('.xlsx'):
            source_file += '.xlsx'

        # Baca data
        df = pd.read_excel(source_file, sheet_name=0)  # Membaca sheet pertama

        # Validasi kolom
        req_cols = ['Nama Pelanggan', 'Tgl. Faktur', 'No. Pelanggan']
        missing = [col for col in req_cols if col not in df.columns]
        if missing:
            raise ValueError(Fore.RED + f"❌ Kolom {missing} tidak ditemukan dalam file sumber!" + Style.RESET_ALL)

        # Meminta lokasi
        print(Fore.YELLOW + "Pilih lokasi:" + Style.RESET_ALL)
        print("1. Surabaya\n2. Semarang\n3. Samarinda\n4. Bagong Jaya")
        location = input(Fore.GREEN + "Masukkan nomor: " + Style.RESET_ALL).strip()

        # Validasi lokasi
        id_tku_map = {
            '1': '0947793543518000000000',
            '2': '0947793543518000000001', 
            '3': '0947793543518000000002',
            '4': '0712982594609000000000'
        }
        id_tku = id_tku_map.get(location, '')
        if not id_tku:
            raise ValueError(Fore.RED + "Lokasi tidak valid!" + Style.RESET_ALL)

        # Load workbook dari file template
        wb = load_workbook(template_file)
        sheet = wb['Faktur']

        # Proses data
        current_row = 4
        for _, row in df.iterrows():
            if pd.isna(row['Nama Pelanggan']) or pd.isna(row['Tgl. Faktur']):
                continue
            
            # Format tanggal
            try:
                tgl = row['Tgl. Faktur'].strftime('%d/%m/%Y')
            except:
                tgl = 'TANGGAL_INVALID'
            
            # Isi data
            sheet[f'A{current_row}'] = current_row - 3
            sheet[f'B{current_row}'] = tgl
            sheet[f'C{current_row}'] = 'Normal'
            sheet[f'D{current_row}'] = '04'
            sheet[f'R{current_row}'] = row['No. Pelanggan']
            sheet[f'N{current_row}'] = row['Nama Pelanggan']
            sheet[f'I{current_row}'] = id_tku
            
            if use_ref:
                sheet[f'G{current_row}'] = row.get('No. Faktur', '')
                
            current_row += 1
        
        wb.save(template_file)
        print(Fore.GREEN + f"\n✅ Sukses! Total data: {current_row-4}" + Style.RESET_ALL)
        
    except Exception as e:
        print(Fore.RED + f"❌ Error: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":
    # Contoh pemanggilan
    process_customer("template.xlsx", "DataCustomer.xlsx", True)
