# goods.py
from openpyxl import load_workbook
import pandas as pd

def process_goods():
    print("\n=== INPUT DATA BARANG ===")
    try:
        # Input pengguna
        template = input("Nama file template (contoh: FF.xlsx): ").strip()
        source = input("Nama file sumber barang (contoh: DataBarang.xlsx): ").strip()
        
        # Baca data
        wb = load_workbook(template)
        sheet = wb['DetailFaktur']
        
        df = pd.read_excel(source)
        
        # Validasi kolom
        req_cols = ['Nama Barang', 'Harga DPP', 'Qty']
        missing = [col for col in req_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Kolom {missing} tidak ditemukan!")
        
        # Proses data
        current_row = 2
        for _, row in df.iterrows():
            if pd.isna(row['Nama Barang']):
                continue
                
            # Hitung nilai
            harga_satuan = round(row['Harga DPP'], 2)
            qty = row['Qty']
            dpp = round(harga_satuan * qty, 2)
            dpp_nilai_lain = round(dpp * (11/12), 2)
            ppn = round(dpp_nilai_lain * 0.12)
            
            # Isi data ke Excel
            sheet[f'A{current_row}'] = current_row - 1                 # Baris
            sheet[f'B{current_row}'] = 'A'                             # Barang/Jasa
            sheet[f'C{current_row}'] = row.get('Kode Barang', '000000')# Kode Barang
            sheet[f'D{current_row}'] = row['Nama Barang']              # Nama Barang
            sheet[f'E{current_row}'] = 'UM.0002'                       # Satuan Ukur
            sheet[f'F{current_row}'] = harga_satuan                    # Harga Satuan
            sheet[f'G{current_row}'] = qty                             # Jumlah Barang
            sheet[f'H{current_row}'] = 0                               # Total Diskon
            sheet[f'I{current_row}'] = dpp                             # DPP
            sheet[f'J{current_row}'] = dpp_nilai_lain                  # DPP Nilai Lain
            sheet[f'K{current_row}'] = 12                              # Tarif PPN (%)
            sheet[f'L{current_row}'] = ppn                             # PPN
            sheet[f'M{current_row}'] = 0                               # Tarif PPnBM
            sheet[f'N{current_row}'] = 0                               # PPnBM
            
            current_row += 1
        
        wb.save(template)
        print(f"\n✅ Sukses! Total data: {current_row-2}")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")

if __name__ == "__main__":
    process_goods()
