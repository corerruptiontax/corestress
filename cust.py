# cust.py
from openpyxl import load_workbook
import pandas as pd
from colorama import Fore, Style
from tqdm import tqdm
from datetime import datetime

def process_customer(template_file, source_file, use_ref, id_tku):
    print(Fore.CYAN + "\n=== PROSES DATA CUSTOMER ===" + Style.RESET_ALL)
    try:
        # Validasi file
        if not source_file.endswith('.xlsx'):
            source_file += '.xlsx'

        # Baca data
        df = pd.read_excel(source_file, sheet_name=0)
        
        # Validasi kolom wajib
        req_cols = ['Nama Pelanggan', 'Tgl. Faktur', 'No. Pelanggan']
        missing = [col for col in req_cols if col not in df.columns]
        if missing:
            raise ValueError(f"Kolom {missing} tidak ditemukan!")

        # Load workbook
        wb = load_workbook(template_file)
        sheet = wb['Faktur']

        # Konfigurasi bulan Indonesia
        bulan_translation = {
            'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr',
            'Mei': 'May', 'Jun': 'Jun', 'Jul': 'Jul', 'Agu': 'Aug',
            'Sep': 'Sep', 'Okt': 'Oct', 'Nov': 'Nov', 'Des': 'Dec'
        }

        # Proses data dengan progress bar
        current_row = 4
        for _, row in tqdm(df.iterrows(), desc="Memproses data", unit="row"):
            if pd.isna(row['Nama Pelanggan']) or pd.isna(row['Tgl. Faktur']):
                continue

            tgl_value = row['Tgl. Faktur']
            tgl = "INVALID_DATE"
            
            try:
                # =============================================
                # BEGIN: LOGIC KONVERSI TANGGAL TERUPDATE
                # =============================================
                
                # 1. Handle datetime object dari Excel
                if isinstance(tgl_value, pd.Timestamp):
                    tgl = tgl_value.strftime("%d/%m/%Y")
                
                # 2. Handle string
                else:
                    tgl_input = str(tgl_value).strip()
                    
                    # Coba format ISO (YYYY-MM-DD dengan/tanpa waktu)
                    try:
                        date_part = tgl_input.split()[0]  # Ambil bagian tanggal
                        parsed_date = datetime.strptime(date_part, "%Y-%m-%d")
                        tgl = parsed_date.strftime("%d/%m/%Y")
                    except:
                        # Normalisasi delimiter dan terjemahkan bulan
                        normalized = tgl_input.replace('-', ' ').replace('/', ' ')
                        for id_bulan, en_bulan in bulan_translation.items():
                            normalized = normalized.replace(id_bulan, en_bulan)
                        
                        # Coba format DD MMM YYYY
                        try:
                            parsed_date = datetime.strptime(normalized, "%d %b %Y")
                            tgl = parsed_date.strftime("%d/%m/%Y")
                        except:
                            # Coba format DD/MM/YYYY asli
                            try:
                                parsed_date = datetime.strptime(normalized, "%d %m %Y")
                                tgl = parsed_date.strftime("%d/%m/%Y")
                            except:
                                raise

                # END: LOGIC KONVERSI
                # =============================================

            except Exception as e:
                print(Fore.YELLOW + f"⚠️ Gagal konversi: '{tgl_value}'" + Style.RESET_ALL)

            # Tulis data ke Excel
            sheet[f'A{current_row}'] = current_row - 3
            sheet[f'B{current_row}'] = tgl
            sheet[f'C{current_row}'] = 'Normal'
            sheet[f'D{current_row}'] = '04'
            sheet[f'L{current_row}'] = 'IDN'
            sheet[f'R{current_row}'] = row['No. Pelanggan']
            sheet[f'N{current_row}'] = row['Nama Pelanggan']
            sheet[f'I{current_row}'] = id_tku
            
            if use_ref:
                sheet[f'G{current_row}'] = row.get('No. Faktur', '')
                
            current_row += 1

        # Simpan dan tutup
        wb.save(template_file)
        print(Fore.GREEN + f"✅ Sukses! Total data: {current_row-4}" + Style.RESET_ALL)

    except Exception as e:
        print(Fore.RED + f"❌ Error: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":
    process_customer("template.xlsx", "data.xlsx", True, "ID_TKU_TEST")
