import logging
import sys
import os  # Tambahkan impor os untuk membuat folder
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

# Fungsi untuk log info (hanya ditulis ke file log.txt)
def log_info(message):
    logger.info(message)

# Fungsi untuk log peringatan (ditulis ke file dan terminal, berwarna kuning tanpa emoticon di terminal)
def log_warning(message):
    # Pesan untuk terminal tanpa emoticon
    terminal_message = message
    colored_message = f"{Fore.YELLOW}{terminal_message}{Style.RESET_ALL}"
    logger.warning(colored_message)

# Fungsi untuk log penting (ditulis ke file dan terminal, tanpa emoticon di terminal)
def log_important(message):
    temp_handler = logging.StreamHandler(sys.stdout)
    temp_handler.setLevel(logging.INFO)
    temp_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(temp_handler)
    logger.info(message)
    logger.removeHandler(temp_handler)