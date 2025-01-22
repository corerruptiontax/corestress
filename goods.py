import pandas as pd  
from openpyxl import load_workbook  
import argparse  # Import argparse untuk menangani argumen  
  
def populate_detail_faktur(ff_file, source_file):  
    # Load workbook dari file tujuan  
    wb = load_workbook(ff_file)  
    detail_faktur_sheet = wb['DetailFaktur']  
  
    # Membaca data dari file sumber        
    try:        
        source_data = pd.read_excel(source_file)        
    except FileNotFoundError:        
        print("File tidak ditemukan. Pastikan nama file sudah benar.")        
        return  
  
    # Mengisi header  
    headers = [  
        "Baris", "Barang/Jasa", "Kode Barang", "Nama Barang", "Nama Satuan Ukur",  
        "Harga Satuan", "Jumlah Barang Jasa", "Total Diskon", "DPP", "DPP Nilai Lain",  
        "Tarif PPN", "PPN", "Tarif PPnBM", "PPnBM"  
    ]  
  
    for col_num, header in enumerate(headers, 1):  
        detail_faktur_sheet.cell(row=1, column=col_num, value=header)  
  
    # Mengisi data  
    current_row = 2  # Mulai dari baris kedua  
    for index, row in source_data.iterrows():  
        if pd.notna(row.get('Nama Barang')) and row['Nama Barang'] != "":  
            detail_faktur_sheet.cell(row=current_row, column=1, value=row.get('Baris', ''))  
            detail_faktur_sheet.cell(row=current_row, column=2, value=row.get('Barang/Jasa', 'A'))  # Mengisi Barang/Jasa  
            detail_faktur_sheet.cell(row=current_row, column=3, value=row.get('Kode Barang', '000000'))  # Mengisi Kode Barang  
            detail_faktur_sheet.cell(row=current_row, column=4, value=row.get('Nama Barang', ''))  
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
            ppn = round(dpp_nilai_lain * (tarif_ppn / 100))  # Menghilangkan desimal  
            detail_faktur_sheet.cell(row=current_row, column=11, value=tarif_ppn)  
            detail_faktur_sheet.cell(row=current_row, column=12, value=ppn)  # PPN dibulatkan  
            detail_faktur_sheet.cell(row=current_row, column=13, value=0)  # Tarif PPnBM  
            detail_faktur_sheet.cell(row=current_row, column=14, value=0)  # PPnBM  
  
            current_row += 1  
  
    # Mengatur format kolom  
    detail_faktur_sheet.column_dimensions['A'].number_format = 'General'  
    detail_faktur_sheet.column_dimensions['B'].number_format = 'General'  
    detail_faktur_sheet.column_dimensions['C'].number_format = '@'  # Text  
    detail_faktur_sheet.column_dimensions['D'].number_format = 'General'  
    detail_faktur_sheet.column_dimensions['E'].number_format = 'General'  
    detail_faktur_sheet.column_dimensions['F'].number_format = '0.00'  # Decimal 2 angka  
    detail_faktur_sheet.column_dimensions['G'].number_format = '0.00'  # Decimal 2 angka  
    detail_faktur_sheet.column_dimensions['H'].number_format = '0.00'  # Decimal 2 angka  
    detail_faktur_sheet.column_dimensions['I'].number_format = '0.00'  # Decimal 2 angka  
    detail_faktur_sheet.column_dimensions['J'].number_format = '0.00'  # Decimal 2 angka  
    detail_faktur_sheet.column_dimensions['K'].number_format = '0'  # Number  
    detail_faktur_sheet.column_dimensions['L'].number_format = '0.00'  # Decimal 2 angka  
    detail_faktur_sheet.column_dimensions['M'].number_format = 'General'  
    detail_faktur_sheet.column_dimensions['N'].number_format = '0.00'  # Decimal 2 angka  
  
    # Menyimpan perubahan  
    wb.save(ff_file)  
  
    # Mengambil nilai terakhir di kolom "Baris"  
    last_row_value = detail_faktur_sheet.cell(row=current_row - 1, column=1).value  
  
    # Menghitung total baris di sheet  
    total_baris = current_row - 1  # Total baris yang diisi  
  
    print(f"Data berhasil dipindahkan ke sheet 'DetailFaktur'. Total Baris: {total_baris}, Angka terakhir di kolom 'Baris': {last_row_value}")  
  
if __name__ == "__main__":  
    # Menggunakan argparse untuk menangani argumen  
    parser = argparse.ArgumentParser(description='Proses data dari file sumber ke file tujuan.')  
    parser.add_argument('source_file', type=str, help='Nama file sumber (misal: "Faktur Pajak Desember 2024.xlsx")')  
    parser.add_argument('ff_file', type=str, help='Nama file tujuan (misal: "GG.xlsx")')  
  
    args = parser.parse_args()  
    populate_detail_faktur(args.ff_file, args.source_file)  
