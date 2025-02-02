# db.py
from openpyxl import load_workbook
import pandas as pd
from colorama import Fore, Style
from tqdm import tqdm
import os

def full_vlookup(template_file, loc_data):
    """
    Melakukan VLOOKUP data dari file KTP ke template
    
    Parameter:
    template_file (str): Path file template Excel
    loc_data (dict): Data konfigurasi lokasi
    """
    print(Fore.CYAN + "\n=== PROSES VLOOKUP ===" + Style.RESET_ALL)
    try:
        # Input file KTP
        ktp_file = input(Fore.GREEN + "Nama file KTP (contoh: KTP.xlsx): " + Style.RESET_ALL).strip()
        if not ktp_file.endswith('.xlsx'):
            ktp_file += '.xlsx'
            
        if not os.path.exists(ktp_file):
            raise FileNotFoundError(f"File KTP '{ktp_file}' tidak ditemukan!")

        # Validasi sheet KTP
        ktp_sheet_name = loc_data['ktp_sheet']
        if ktp_sheet_name not in pd.ExcelFile(ktp_file).sheet_names:
            raise ValueError(f"Sheet '{ktp_sheet_name}' tidak ada di file KTP!")

        # Buka template
        print(Fore.BLUE + "📂 Membuka template Excel..." + Style.RESET_ALL)
        wb = load_workbook(template_file)
        sheet = wb['Faktur']
        
        # Ambil kode BBN (kolom R) dan skip baris END
        kode_bbn = []
        for row in sheet.iter_rows(min_row=4, min_col=18, max_col=18):  # Kolom R = 18
            cell_value = row[0].value
            if cell_value is not None and str(cell_value).strip().upper() != "END":
                kode_bbn.append(cell_value)
        
        # Baca data KTP
        print(Fore.BLUE + "🔍 Membaca data KTP..." + Style.RESET_ALL)
        ktp_df = pd.read_excel(
            ktp_file,
            sheet_name=ktp_sheet_name,
            usecols="C,D,E,H,I,J,K",
            header=0,
            dtype=str
        )
        
        # Proses mapping
        updated = 0
        total_valid = len(kode_bbn)
        
        for idx, kode in tqdm(enumerate(kode_bbn, 4), total=total_valid, desc="VLOOKUP"):
            result = ktp_df[ktp_df.iloc[:, 0] == str(kode)]
            if not result.empty:
                data = result.iloc[0]
                sheet[f'S{idx}'] = data.iloc[0]
                sheet[f'K{idx}'] = data.iloc[1]
                sheet[f'M{idx}'] = data.iloc[2]
                sheet[f'J{idx}'] = data.iloc[3]
                sheet[f'Q{idx}'] = data.iloc[4]
                sheet[f'N{idx}'] = data.iloc[5]
                sheet[f'O{idx}'] = data.iloc[6]
                updated += 1
        
        # Simpan file
        output_file = template_file.replace(".xlsx", "_FULL.xlsx")
        print(Fore.BLUE + "💾 Menyimpan hasil..." + Style.RESET_ALL)
        wb.save(output_file)
        
        print(Fore.GREEN + f"✅ Diupdate: {updated}/{total_valid}" + Style.RESET_ALL)
        print(Fore.BLUE + f"📁 File hasil: {output_file}" + Style.RESET_ALL)
        
    except Exception as e:
        error_msg = f"""
        ❌ ERROR VLOOKUP:
        File: {os.path.basename(ktp_file) if 'ktp_file' in locals() else 'N/A'}
        Detil Error: {str(e)}
        """
        print(Fore.RED + error_msg + Style.RESET_ALL)
