import logging
import sys
import os
from colorama import init, Fore, Style

# Inisialisasi colorama
init(autoreset=True)

# Buat folder Trans Logs jika belum ada
trans_logs_folder = "Trans Logs"
os.makedirs(trans_logs_folder, exist_ok=True)

# Setup logger untuk file log.txt di folder Trans Logs (semua level, menggunakan encoding utf-8)
log_file_path = os.path.join(trans_logs_folder, 'log.txt')
file_handler = logging.FileHandler(log_file_path, mode='w', encoding='utf-8')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Setup logger untuk terminal (hanya WARNING ke atas, dan pesan dari log_important)
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.WARNING)
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Setup logger
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.handlers = []  # Hapus handler default
logger.addHandler(file_handler)  # Tambahkan handler untuk file
logger.addHandler(console_handler)  # Tambahkan handler untuk terminal

# Fungsi untuk log info (hanya ditulis ke file log.txt, tidak ke terminal)
def log_info(message):
    logger.info(message)  # Ke file tanpa emoticon

# Fungsi untuk log peringatan (ditulis ke file dan terminal, kuning + emoticon di terminal)
def log_warning(message):
    logger.warning(message)  # Ke file tanpa emoticon
    console_message = f"{Fore.YELLOW}⚠️ {message}{Style.RESET_ALL}"
    print(console_message)  # Ke terminal dengan warna dan emoticon

# Fungsi untuk log penting (ditulis ke file dan terminal, putih + emoticon di terminal)
def log_important(message):
    logger.info(message)  # Ke file tanpa emoticon
    console_message = f"{Fore.WHITE}📢 {message}{Style.RESET_ALL}"
    print(console_message)  # Ke terminal dengan warna dan emoticon

# Fungsi untuk log sukses (ditulis ke file dan terminal, hijau + emoticon di terminal)
def log_success(message):
    logger.info(message)  # Ke file tanpa emoticon
    console_message = f"{Fore.GREEN}✅ {message}{Style.RESET_ALL}"
    print(console_message)  # Ke terminal dengan warna dan emoticon

# Fungsi untuk log error (ditulis ke file dan terminal, merah + emoticon di terminal)
def log_error(message):
    logger.error(message)  # Ke file tanpa emoticon
    console_message = f"{Fore.RED}❌ {message}{Style.RESET_ALL}"
    print(console_message)  # Ke terminal dengan warna dan emoticon

# Fungsi untuk log ringkasan (ditulis ke file dan terminal, magenta + emoticon di terminal)
def log_summary(message):
    logger.info(message)  # Ke file tanpa emoticon
    console_message = f"{Fore.MAGENTA}📊 {message}{Style.RESET_ALL}"
    print(console_message)  # Ke terminal dengan warna dan emoticon