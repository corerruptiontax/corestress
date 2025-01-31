# goods.py
import pandas as pd  
from openpyxl import load_workbook  
from colorama import Fore, Style
from tqdm import tqdm

def populate_detail_faktur(template_file, source_file):  
    print(Fore.CYAN + "\n=== PROSES DATA BARANG ===" + Style.RESET_ALL)
    try:
        if not source_file.endswith('.xlsx'):
            source_file += '.xlsx'

        wb = load_workbook(template_file)  
        sheet = wb['DetailFaktur']
        
        if sheet.max_row > 1:
            sheet.delete_rows(2, sheet.max_row)

        source_data = pd.read_excel(source_file)
        required_cols = ['Nama Barang', 'Harga DPP', 'Qty']
        missing = [col for col in required_cols if col not in source_data.columns]
        if missing:
            raise ValueError(f"Kolom {missing} tidak ditemukan!")

        current_row = 2
        for _, row in tqdm(source_data.iterrows(), desc="Memproses barang", unit="row"):
            if pd.notna(row.get('Nama Barang')) and row['Nama Barang'] != "":
                sheet.cell(row=current_row, column=1, value=row.get('Baris', ''))  
                sheet.cell(row=current_row, column=2, value=row.get('Barang/Jasa', 'A'))  
                sheet.cell(row=current_row, column=3, value=row.get('Kode Barang', '000000'))  
                sheet.cell(row=current_row, column=4, value=row['Nama Barang'])  
                sheet.cell(row=current_row, column=5, value=row.get('Nama Satuan Ukur', 'UM.0002'))  

                harga_satuan = round(row.get('Harga DPP', 0), 2)  
                jumlah_barang = row.get('Qty', 0)  
                dpp = round(harga_satuan * jumlah_barang, 2)  
                dpp_nilai_lain = round(dpp * (11 / 12), 2)  
                ppn = round(dpp_nilai_lain * 0.12, 2)  

                sheet.cell(row=current_row, column=6, value=harga_satuan)  
                sheet.cell(row=current_row, column=7, value=jumlah_barang)  
                sheet.cell(row=current_row, column=8, value=0)  
                sheet.cell(row=current_row, column=9, value=dpp)  
                sheet.cell(row=current_row, column=10, value=dpp_nilai_lain)  
                sheet.cell(row=current_row, column=11, value=12)  
                sheet.cell(row=current_row, column=12, value=ppn)  
                sheet.cell(row=current_row, column=13, value=0)  
                sheet.cell(row=current_row, column=14, value=0)  

                current_row += 1

        wb.save(template_file)
        print(Fore.GREEN + f"✅ Berhasil diproses: {current_row-2} baris" + Style.RESET_ALL)

    except Exception as e:
        print(Fore.RED + f"❌ Error: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":  
    populate_detail_faktur("template.xlsx", "data.xlsx")
