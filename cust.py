import pandas as pd          
from openpyxl import load_workbook          
import argparse        
  
def populate_faktur(ff_file, source_file, use_referensi, location):          
    # Mapping lokasi ke ID TKU      
    id_tku_mapping = {      
        'Surabaya': '0947793543518000000000',      
        'Semarang': '0947793543518000000001',      
        'Samarinda': '0947793543518000000002',      
        'Bagong Jaya': '0712982594609000000000'      
    }      
        
    id_tku = id_tku_mapping.get(location)  # Mendapatkan ID TKU berdasarkan lokasi  
  
    # Load workbook dari file tujuan  
    wb = load_workbook(ff_file)  
    faktur_sheet = wb['Faktur']  
  
    # Menambahkan header di baris ketiga  
    headers = [  
        "Baris", "Tanggal Faktur", "Jenis Faktur", "Kode Transaksi",  
        "Keterangan Tambahan", "Dokumen Pendukung", "Referensi", "Cap Fasilitas",  
        "ID TKU Penjual", "NPWP/NIK Pembeli", "Jenis ID Pembeli", "Negara Pembeli",  
        "Nomor Dokumen Pembeli", "Nama Pembeli", "Alamat Pembeli", "Email Pembeli",  
        "ID TKU Pembeli", "Kode Pelanggan"  
    ]  
  
    # Mengisi header di baris ketiga  
    for col_num, header in enumerate(headers, start=1):  
        faktur_sheet.cell(row=3, column=col_num, value=header)  
  
    current_row = 4  # Mulai dari baris keempat untuk data  
    total_baris = 0  # Inisialisasi total baris yang dipindahkan  
  
    # Membaca data dari file sumber        
    try:        
        source_data = pd.read_excel(source_file)        
    except FileNotFoundError:        
        print("File tidak ditemukan. Pastikan nama file sudah benar.")        
        return    
  
    for index, row in source_data.iterrows():  
        # Pastikan kolom yang relevan tidak kosong    
        if pd.notna(row['Nama Pelanggan']) and row['Nama Pelanggan'] != "" and pd.notna(row['Tgl. Faktur']):  
            faktur_sheet.cell(row=current_row, column=1, value=int(row['Baris']))  # Baris sebagai integer  
              
            # Ambil tanggal dan format ke DD/MM/YYYY        
            tanggal_faktur = pd.to_datetime(row['Tgl. Faktur'], errors='coerce')        
            if pd.notna(tanggal_faktur):        
                faktur_sheet.cell(row=current_row, column=2, value=tanggal_faktur.strftime('%d/%m/%Y'))  # Tanggal Faktur        
            else:        
                faktur_sheet.cell(row=current_row, column=2, value='')  # Kosongkan jika tidak valid        
  
            faktur_sheet.cell(row=current_row, column=3, value='Normal')  # Jenis Faktur        
            faktur_sheet.cell(row=current_row, column=4, value='04')  # Kode Transaksi        
            faktur_sheet.cell(row=current_row, column=5, value='')  # Keterangan Tambahan        
            faktur_sheet.cell(row=current_row, column=6, value='')  # Dokumen Pendukung        
  
            # Mengisi kolom Referensi berdasarkan pilihan      
            if use_referensi:      
                faktur_sheet.cell(row=current_row, column=7, value=row['No. Faktur'])  # Referensi      
            else:      
                faktur_sheet.cell(row=current_row, column=7, value='')  # Kosongkan jika tidak menggunakan referensi      
  
            faktur_sheet.cell(row=current_row, column=8, value='')  # Cap Fasilitas (dibiarkan kosong)    
            faktur_sheet.cell(row=current_row, column=9, value=id_tku)  # ID TKU Penjual            
            faktur_sheet.cell(row=current_row, column=10, value='')  # NPWP/NIK Pembeli            
            faktur_sheet.cell(row=current_row, column=11, value='')  # Jenis ID Pembeli            
            faktur_sheet.cell(row=current_row, column=12, value='IDN')  # Negara Pembeli            
            faktur_sheet.cell(row=current_row, column=13, value='')  # Nomor Dokumen Pembeli            
            faktur_sheet.cell(row=current_row, column=14, value=row['Nama Pelanggan'])  # Nama Pembeli          
            faktur_sheet.cell(row=current_row, column=15, value='')  # Alamat Pembeli            
            faktur_sheet.cell(row=current_row, column=16, value='')  # Email Pembeli           
            faktur_sheet.cell(row=current_row, column=17, value='')  # ID TKU Pembeli (dibiarkan kosong)  
            faktur_sheet.cell(row=current_row, column=18, value=row['No. Pelanggan'])  # Kode Pelanggan          
  
            current_row += 1          
            total_baris += 1  # Increment total baris yang dipindahkan        
  
    # Menyimpan perubahan  
    wb.save(ff_file)  
    print(f"Data berhasil dipindahkan ke sheet 'Faktur'. Total Baris: {total_baris}")  
  
if __name__ == "__main__":  
    # Menggunakan argparse untuk menangani argumen  
    parser = argparse.ArgumentParser(description='Proses data dari file sumber ke file tujuan.')  
    parser.add_argument('source_file', type=str, help='Nama file sumber (misal: "Faktur Pajak Desember 2024.xlsx")')  
    parser.add_argument('ff_file', type=str, help='Nama file tujuan (misal: "DD.xlsx")')  
    parser.add_argument('--use_referensi', action='store_true', help='Gunakan referensi dari file sumber')  
    parser.add_argument('--location', type=str, required=True, help='Lokasi untuk ID TKU')  
  
    args = parser.parse_args()  
    populate_faktur(args.ff_file, args.source_file, args.use_referensi, args.location)  
