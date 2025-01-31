# db.py
from openpyxl import load_workbook
import pandas as pd
from colorama import Fore, Style

def full_vlookup(template_file):
    print(Fore.CYAN + "\n=== INPUT DATA VLOOKUP ===" + Style.RESET_ALL)
    try:
        ktp_file = input(Fore.GREEN + "Nama file KTP - List (NEW).xlsx (contoh: KTP - List (NEW).xlsx): " + Style.RESET_ALL).strip()
        if not ktp_file.endswith('.xlsx'):
            ktp_file += '.xlsx'
        
        location = get_location()  # Pindahkan logika lokasi ke fungsi terpisah
        
        # Validasi lokasi
        id_tku_map = {
            '1': '0947793543518000000000',
            '2': '0947793543518000000001', 
            '3': '0947793543518000000002',
            '4': '0712982594609000000000'
        }
        id_tku = id_tku_map.get(location, '')
        if not id_tku:
            raise ValueError(Fore.RED + "❌ Lokasi tidak valid!" + Style.RESET_ALL)
        
        wb = load_workbook(template_file)
        sheet = wb['Faktur']
        
        kode_bbn = [cell[0].value for cell in sheet.iter_rows(min_row=4, min_col=18, max_col=18)]
        
        ktp_sheet_name = {
            '1': 'NPWPKTP BBN SBY-BJM (NEW)',
            '2': 'NPWPKTP BBN SMG (NEW)',
            '3': 'NPWPKTP BBN SMD-BPP (NEW)',
            '4': 'NPWPKTP BJ (NEW)'
        }[location]
        
        ktp_df = pd.read_excel(ktp_file, sheet_name=ktp_sheet_name, usecols="C,D,E,H,I,J,K", header=0, dtype=str)
        
        updated = 0
        for idx, kode in enumerate(kode_bbn, 4):
            if pd.isna(kode) or kode == "":
                continue
            
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
        
        output_file = template_file.replace(".xlsx", "_FULL_UPDATED.xlsx")
        wb.save(output_file)
        print(Fore.GREEN + f"Total Baris Diupdate: {updated}/{len(kode_bbn)}" + Style.RESET_ALL)
        print(Fore.BLUE + f"File hasil: {output_file} 📁" + Style.RESET_ALL)
        
    except Exception as e:
        print(Fore.RED + f"\n❌ Error: {str(e)}" + Style.RESET_ALL)

def get_location():
    print(Fore.YELLOW + "Pilih lokasi:" + Style.RESET_ALL)
    print("1. Surabaya\n2. Semarang\n3. Samarinda\n4. Bagong Jaya")
    location = input(Fore.GREEN + "Masukkan nomor: " + Style.RESET_ALL).strip()
    if location not in ['1', '2', '3', '4']:
        raise ValueError(Fore.RED + "❌ Lokasi tidak valid! Silakan masukkan nomor 1-4." + Style.RESET_ALL)
    return location

if __name__ == "__main__":
    full_vlookup("template.xlsx")  # Contoh pemanggilan
