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

def display_welcome_message():
    # Membaca isi dari file ascii_art.txt
    try:
        with open("ascii_art.txt", "r", encoding="utf-8") as file:
            ascii_art = file.read()
            print(ascii_art)  # Menampilkan ASCII art
            time.sleep(3)  # Memberikan jeda selama 3 detik
    except FileNotFoundError:
        print("File ascii_art.txt tidak ditemukan.")
        time.sleep(3)  # Memberikan jeda selama 3 detik jika file tidak ditemukan

def main():
    # Fungsi utama program Anda
    print("")

if __name__ == "__main__":
    display_welcome_message()  # Menampilkan pesan sambutan
    main()  # Menjalankan fungsi utama

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
        
        # Input nama file output
        output_file = input(Fore.GREEN + "Masukkan nama file output (tanpa .xlsx): " + Style.RESET_ALL).strip()
        
        # Step 1: Buat template
        print(Fore.YELLOW + "\n=== STEP 1: BUAT TEMPLATE ===" + Style.RESET_ALL)
        create_template(output_file, loc_data)
        
        # Input file sumber
        source_file = input(Fore.GREEN + "Masukkan nama file sumber data: " + Style.RESET_ALL).strip()
        if not source_file.endswith('.xlsx'):
            source_file += '.xlsx'

        # Step 2: Proses customer
        print(Fore.YELLOW + "\n=== STEP 2: INPUT DATA CUSTOMER ===" + Style.RESET_ALL)
        use_ref = input(Fore.GREEN + "Gunakan referensi? (y/n): ").lower() == 'y'
        process_customer(f"{output_file}.xlsx", source_file, use_ref, loc_data['id_tku'])
        
        # Step 3: Proses barang
        print(Fore.YELLOW + "\n=== STEP 3: INPUT DATA BARANG ===" + Style.RESET_ALL)
        populate_detail_faktur(f"{output_file}.xlsx", source_file)
        
        # Step 4: VLOOKUP
        print(Fore.YELLOW + "\n=== STEP 4: VLOOKUP DATA ===" + Style.RESET_ALL)
        full_vlookup(f"{output_file}.xlsx", loc_data)
        
        print(Fore.GREEN + "\n✅ SEMUA PROSES SELESAI!" + Style.RESET_ALL)
        print(Fore.BLUE + f"📁 File output: {output_file}.xlsx" + Style.RESET_ALL)

    except Exception as e:
        print(Fore.RED + f"\n❌ ERROR UTAMA: {str(e)}" + Style.RESET_ALL)

if __name__ == "__main__":
    main()
