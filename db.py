# db.py
import pandas as pd
from openpyxl import load_workbook

def full_vlookup():
    print("\n=== INPUT DATA VLOOKUP ===")
    try:
        # Input pengguna
        ff_file = input("Nama file FF.xlsx (contoh: FF.xlsx): ").strip()
        ktp_file = input("Nama file KTP - List (NEW).xlsx (contoh: KTP - List (NEW).xlsx): ").strip()
        
        print("Pilih lokasi:")
        print("1. Surabaya\n2. Semarang\n3. Samarinda\n4. Bagong Jaya")
        location = input("Masukkan nomor: ").strip()
        
        # Validasi lokasi
        id_tku_map = {
            '1': '0947793543518000000000',
            '2': '0947793543518000000001', 
            '3': '0947793543518000000002',
            '4': '0712982594609000000000'
        }
        id_tku = id_tku_map.get(location, '')
        if not id_tku:
            raise ValueError("Lokasi tidak valid!")
        
        # Buka file FF.xlsx
        wb = load_workbook(ff_file)
        sheet = wb['Faktur']
        
        # Ambil SEMUA kode pelanggan dari kolom R
        kode_bbn = [cell[0].value for cell in sheet.iter_rows(min_row=4, min_col=18, max_col=18)]
        
        # Baca KTP List
        ktp_sheet_name = {
            '1': 'NPWPKTP BBN SBY-BJM (NEW)',
            '2': 'NPWPKTP BBN SMG (NEW)',
            '3': 'NPWPKTP BBN SMD-BPP (NEW)',
            '4': 'NPWPKTP BJ (NEW)'
        }[location]
        
        ktp_df = pd.read_excel(
            ktp_file,
            sheet_name=ktp_sheet_name,
            usecols="C,D,E,H,I,J,K",  # Kolom C-K
            header=0,
            dtype=str
        )
        
        # Proses lookup untuk SEMUA baris
        updated = 0
        for idx, kode in enumerate(kode_bbn, 4):
            if pd.isna(kode) or kode == "":
                continue
                
            # Cari exact match
            result = ktp_df[ktp_df.iloc[:, 0] == str(kode)]
            
            if not result.empty:
                data = result.iloc[0]
                sheet[f'S{idx}'] = data.iloc[0]  # Kode Pelanggan KTP List
                sheet[f'K{idx}'] = data.iloc[1]  # Jenis ID Pembeli
                sheet[f'M{idx}'] = data.iloc[2]  # Nomor Dokumen Pembeli
                sheet[f'J{idx}'] = data.iloc[3]  # NPWP/NIK Pembeli
                sheet[f'Q{idx}'] = data.iloc[4]  # ID TKU Pembeli
                sheet[f'N{idx}'] = data.iloc[5]  # Nama Pembeli
                sheet[f'O{idx}'] = data.iloc[6]  # Alamat Pembeli
                updated += 1
        
        # Simpan ke file baru
        output_file = ff_file.replace(".xlsx", "_FULL_UPDATED.xlsx")
        wb.save(output_file)
        print(f"Total Baris Diupdate: {updated}/{len(kode_bbn)}")
        print(f"File hasil: {output_file}")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")

if __name__ == "__main__":
    full_vlookup()
