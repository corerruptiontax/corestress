import os  # Tambahkan ini
import sys
import time  # Import modul time untuk delay
from pathlib import Path
import cowsay
from colorama import init, Fore, Style

# Tambahkan path ke sys.path agar Python mengenali folder
sys.path.append(str(Path(__file__).parent / "config"))
sys.path.append(str(Path(__file__).parent / "src"))
sys.path.append(str(Path(__file__).parent))

from src.setupcore import create_template
from src.cust import process_customer
from src.goods import populate_detail_faktur
from src.db import full_vlookup

# Inisialisasi colorama untuk warna terminal
init(autoreset=True)

# Konfigurasi lokasi
LOCATION_CONFIG = {
    '1': {
        'name': 'Surabaya',
        'npwp': '0947793543518000',
        'id_tku': '0947793543518000000000',
        'ktp_sheet': 'NPWPKTP BBN SBY-BJM (NEW)'
    },
    '2': {
        'name': 'Semarang',
        'npwp': '0947793543518000',
        'id_tku': '0947793543518000000001',
        'ktp_sheet': 'NPWPKTP BBN SMG (NEW)'
    },
    '3': {
        'name': 'Samarinda',
        'npwp': '0947793543518000',
        'id_tku': '0947793543518000000002',
        'ktp_sheet': 'NPWPKTP BBN SMD-BPP (NEW)'
    },
    '4': {
        'name': 'Bagong Jaya',
        'npwp': '0712982594609000',
        'id_tku': '0712982594609000000000',
        'ktp_sheet': 'NPWPKTP BJ (NEW)'
    }
}

def get_location():
    """Meminta input lokasi dari user"""
    print(Fore.YELLOW + "Pilih lokasi:" + Style.RESET_ALL)
    print("1. Surabaya\n2. Semarang\n3. Samarinda\n4. Bagong Jaya")
    location = input(Fore.GREEN + "Masukkan nomor: " + Style.RESET_ALL).strip()
    if location not in LOCATION_CONFIG:
        raise ValueError(Fore.RED + "Lokasi tidak valid!" + Style.RESET_ALL)
    return LOCATION_CONFIG[location]

def main():
    """Program utama"""
    print(Fore.CYAN + "\n=== MEMULAI PROSES ===" + Style.RESET_ALL)
    
    try:
        # Input lokasi
        loc_data = get_location()
        print(Fore.BLUE + f"📍 Lokasi dipilih: {loc_data['name']}" + Style.RESET_ALL)
        
        # Input nama file output (otomatis menambahkan .xlsx jika tidak diketik)
        output_file = input(Fore.GREEN + "Masukkan nama file output: " + Style.RESET_ALL).strip()
        if not output_file.lower().endswith('.xlsx'):
            output_file += ".xlsx"

        # Tentukan folder template
        template_folder = os.path.join(Path(__file__).parent, "template")

        # Pastikan folder template ada
        if not os.path.exists(template_folder):
            os.makedirs(template_folder)

        # Path lengkap untuk file template
        template_file = os.path.join(template_folder, output_file)

        # Step 1: Buat template
        print(Fore.YELLOW + "\n=== STEP 1: BUAT TEMPLATE ===" + Style.RESET_ALL)
        create_template(template_file, loc_data)

        # Cek apakah file template berhasil dibuat
        if not os.path.exists(template_file):
            raise FileNotFoundError(f"Template file '{template_file}' tidak ditemukan setelah dibuat.")

        # Input file sumber (otomatis menambahkan .xlsx jika tidak diketik)
        source_file = input(Fore.GREEN + "Masukkan nama file sumber data: " + Style.RESET_ALL).strip()
        source_file = f"{source_file}.xlsx" if not source_file.lower().endswith('.xlsx') else source_file

        # Step 2: Proses customer
        print(Fore.YELLOW + "\n=== STEP 2: INPUT DATA CUSTOMER ===" + Style.RESET_ALL)
        use_ref = input(Fore.GREEN + "Gunakan referensi? (y/n): ").lower() == 'y'
        process_customer(template_file, source_file, use_ref, loc_data['id_tku'])

        # Step 3: Proses barang
        print(Fore.YELLOW + "\n=== STEP 3: INPUT DATA BARANG ===" + Style.RESET_ALL)
        populate_detail_faktur(template_file, source_file)

        # Step 4: VLOOKUP
        print(Fore.YELLOW + "\n=== STEP 4: VLOOKUP DATA ===" + Style.RESET_ALL)
        full_vlookup(template_file, loc_data)

        print(Fore.GREEN + "\n✅ SEMUA PROSES SELESAI!" + Style.RESET_ALL)
        print(Fore.BLUE + f"📁 File hasil akhir: {output_file}" + Style.RESET_ALL)

    except Exception as e:
        print(Fore.RED + f"\n❌ ERROR UTAMA: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":
    main()
