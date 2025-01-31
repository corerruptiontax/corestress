# db.py
from openpyxl import load_workbook
import pandas as pd
from colorama import Fore, Style
from tqdm import tqdm

def full_vlookup(template_file, loc_data):
    print(Fore.CYAN + "\n=== PROSES VLOOKUP ===" + Style.RESET_ALL)
    try:
        ktp_file = input(Fore.GREEN + "Nama file KTP (contoh: KTP.xlsx): " + Style.RESET_ALL).strip()
        if not ktp_file.endswith('.xlsx'):
            ktp_file += '.xlsx'

        wb = load_workbook(template_file)
        sheet = wb['Faktur']
        
        kode_bbn = [cell[0].value for cell in sheet.iter_rows(min_row=4, min_col=18, max_col=18)]
        
        ktp_df = pd.read_excel(
            ktp_file,
            sheet_name=loc_data['ktp_sheet'],
            usecols="C,D,E,H,I,J,K",
            header=0,
            dtype=str
        )
        
        updated = 0
        for idx, kode in tqdm(enumerate(kode_bbn, 4), total=len(kode_bbn), desc="VLOOKUP"):
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
        
        output_file = template_file.replace(".xlsx", "_FULL.xlsx")
        wb.save(output_file)
        print(Fore.GREEN + f"✅ Diupdate: {updated}/{len(kode_bbn)}" + Style.RESET_ALL)
        print(Fore.BLUE + f"📁 File hasil: {output_file}" + Style.RESET_ALL)
        
    except Exception as e:
        print(Fore.RED + f"❌ Error: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":
    full_vlookup("template.xlsx", {'ktp_sheet':'TEST'})
