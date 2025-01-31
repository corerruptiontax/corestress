# main.py
from setupcore import create_template
from cust import process_customer
from goods import populate_detail_faktur
from db import full_vlookup
from colorama import init, Fore, Style

# Inisialisasi Colorama
init(autoreset=True)

def main():
    print(Fore.CYAN + "\n=== MEMULAI PROSES ===" + Style.RESET_ALL)
    
    # Meminta nama file output dari pengguna
    output_file_name = input(Fore.GREEN + "Masukkan nama file output (tanpa .xlsx): " + Style.RESET_ALL).strip()
    
    # 1. Setup Core
    print(Fore.YELLOW + "\n=== STEP 1: MENGHASILKAN TEMPLATE EXCEL ===" + Style.RESET_ALL)
    create_template(output_file_name)
    
    # Meminta nama file sumber (sekali saja)
    source_file = input(Fore.GREEN + "Masukkan nama file sumber data (contoh: Data.xlsx): " + Style.RESET_ALL).strip()
    if not source_file.endswith('.xlsx'):
        source_file += '.xlsx'

    # 2. Input Data Customer
    print(Fore.YELLOW + "\n=== STEP 2: INPUT DATA CUSTOMER ===" + Style.RESET_ALL)
    use_ref = input(Fore.GREEN + "Gunakan referensi? (y/n): ").lower() == 'y'  # Menambahkan input untuk use_ref
    process_customer(f"{output_file_name}.xlsx", source_file, use_ref)  # Meneruskan source_file dan use_ref

    # 3. Input Data Barang
    print(Fore.YELLOW + "\n=== STEP 3: INPUT DATA BARANG ===" + Style.RESET_ALL)
    populate_detail_faktur(f"{output_file_name}.xlsx", source_file)  # Meneruskan source_file
    
    # 4. VLOOKUP Data
    print(Fore.YELLOW + "\n=== STEP 4: VLOOKUP DATA ===" + Style.RESET_ALL)
    full_vlookup(f"{output_file_name}.xlsx")  # Memanggil fungsi VLOOKUP dari db.py
    
    print(Fore.GREEN + "\n✅ Semua proses selesai!" + Style.RESET_ALL)

if __name__ == "__main__":
    main()
