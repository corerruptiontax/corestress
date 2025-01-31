# goods.py
import pandas as pd  
from openpyxl import load_workbook  
from colorama import Fore, Style

def populate_detail_faktur(template_file, source_file):  
    print(Fore.CYAN + "\n=== INPUT DATA BARANG ===" + Style.RESET_ALL)
    try:
        # Memastikan file sumber memiliki ekstensi .xlsx
        if not source_file.endswith('.xlsx'):
            source_file += '.xlsx'

        # Load workbook dari file template
        wb = load_workbook(template_file)  
        detail_faktur_sheet = wb['DetailFaktur']  

        # Membaca data dari file sumber
        try:
            source_data = pd.read_excel(source_file)  
        except FileNotFoundError:  
            print(Fore.RED + "❌ File tidak ditemukan. Pastikan nama file sudah benar." + Style.RESET_ALL)  
            return  

        # Mengisi header
        headers = [  
            "Baris", "Barang/Jasa", "Kode Barang", "Nama Barang", "Nama Satuan Ukur",  
            "Harga Satuan", "Jumlah Barang Jasa", "Total Diskon", "DPP", "DPP Nilai Lain",  
            "Tarif PPN", "PPN", "Tarif PPnBM", "PPnBM", "ID TKU Penjual"  
        ]  

        for col_num, header in enumerate(headers, 1):  
            detail_faktur_sheet.cell(row=1, column=col_num, value=header)  

        # Mengisi data
        current_row = 2  # Mulai dari baris kedua  
        for index, row in source_data.iterrows():  
            if pd.notna(row.get('Nama Barang')) and row['Nama Barang'] != "":  
                detail_faktur_sheet.cell(row=current_row, column=1, value=row.get('Baris', ''))  
                detail_faktur_sheet.cell(row=current_row, column=2, value=row.get('Barang/Jasa', 'A'))  
                detail_faktur_sheet.cell(row=current_row, column=3, value=row.get('Kode Barang', '000000'))  
                detail_faktur_sheet.cell(row=current_row, column=4, value=row['Nama Barang'])  
                detail_faktur_sheet.cell(row=current_row, column=5, value=row.get('Nama Satuan Ukur', 'UM.0002'))  

                # Menghitung DPP  
                harga_satuan = round(row.get('Harga DPP', 0), 2)  
                jumlah_barang = row.get('Qty', 0)  
                dpp = round(harga_satuan * jumlah_barang, 2)  
                detail_faktur_sheet.cell(row=current_row, column=6, value=harga_satuan)  
                detail_faktur_sheet.cell(row=current_row, column=7, value=jumlah_barang)  
                detail_faktur_sheet.cell(row=current_row, column=8, value=0)  # Total Diskon  
                detail_faktur_sheet.cell(row=current_row, column=9, value=dpp)  

                # Menghitung DPP Nilai Lain  
                dpp_nilai_lain = round(dpp * (11 / 12), 2)  
                detail_faktur_sheet.cell(row=current_row, column=10, value=dpp_nilai_lain)  

                # Menghitung PPN  
                tarif_ppn = 12  # Misalkan tarif PPN adalah 12%  
                ppn = round(dpp_nilai_lain * (tarif_ppn / 100))  
                detail_faktur_sheet.cell(row=current_row, column=11, value=tarif_ppn)  
                detail_faktur_sheet.cell(row=current_row, column=12, value=ppn)  
                detail_faktur_sheet.cell(row=current_row, column=13, value=0)  # Tarif PPnBM  
                detail_faktur_sheet.cell(row=current_row, column=14, value=0)  # PPnBM  

                current_row += 1  

        # Menyimpan perubahan  
        wb.save(template_file)  

        print(Fore.GREEN + f"Data berhasil dipindahkan ke sheet 'DetailFaktur'." + Style.RESET_ALL)  

    except Exception as e:
        print(Fore.RED + f"❌ Error: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":  
    template_file = input(Fore.GREEN + "Masukkan nama file template (contoh: FF.xlsx): ").strip()
    source_file = input(Fore.GREEN + "Masukkan nama file sumber barang (contoh: DataBarang.xlsx): ").strip()
    
    # Memastikan file sumber memiliki ekstensi .xlsx
    if not source_file.endswith('.xlsx'):
        source_file += '.xlsx'
    
    populate_detail_faktur(template_file, source_file)
