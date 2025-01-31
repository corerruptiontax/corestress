# cust.py
from openpyxl import load_workbook
import pandas as pd

def process_customer():
    print("\n=== INPUT DATA CUSTOMER ===")
    try:
        # Input pengguna
        template = input("Nama file template (contoh: FF.xlsx): ").strip()
        source = input("Nama file sumber customer (contoh: DataCustomer.xlsx): ").strip()
        use_ref = input("Gunakan referensi? (y/n): ").lower() == 'y'
        
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
        
        # Baca data
        wb = load_workbook(template)
        sheet = wb['Faktur']
        
        df = pd.read_excel(
            source,
            parse_dates=['Tgl. Faktur'],
            date_format='mixed'  # Handle berbagai format tanggal
        )
        
        # Validasi kolom
        req_cols = ['Nama Pelanggan', 'Tgl. Faktur', 'No. Pelanggan']
        missing = [col for col in req_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Kolom {missing} tidak ditemukan!")
        
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
            
            if use_ref:
                sheet[f'G{current_row}'] = row.get('No. Faktur', '')
                
            sheet[f'I{current_row}'] = id_tku
            sheet[f'N{current_row}'] = row['Nama Pelanggan']
            
            current_row += 1
        
        wb.save(template)
        print(f"\n✅ Sukses! Total data: {current_row-4}")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")

if __name__ == "__main__":
    process_customer()
