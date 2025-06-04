import os
import sys
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, Optional

import yaml
from colorama import init, Fore, Style

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.append(str(PROJECT_ROOT / "src"))
sys.path.append(str(PROJECT_ROOT / "config"))
sys.path.append(str(PROJECT_ROOT))

from setupcore import create_template
from cust import process_customer
from goods import populate_detail_faktur
from db import full_vlookup

init(autoreset=True)

logging.basicConfig(
    level=logging.WARNING,
    format="%(message)s", # DIUBAH: Hanya pesan saja, tanpa levelname atau timestamp
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

CONFIG_DIR = PROJECT_ROOT / "config"
CONFIG_FILE = CONFIG_DIR / "config.yaml"
TEMPLATE_DIR_NAME = "template"

LOCATION_CONFIG: Dict[str, Dict[str, Any]] = {}

def load_config(config_path: Path) -> Dict[str, Dict[str, Any]]:
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config_data = yaml.safe_load(f)
        if not config_data or 'locations' not in config_data:
            logger.error(Fore.RED + "Format file %s tidak sesuai atau key 'locations' tidak ditemukan.", config_path)
            raise ValueError("Invalid config format.")
        return config_data['locations']
    except FileNotFoundError:
        logger.error(Fore.RED + "❌ ERROR: File konfigurasi %s tidak ditemukan.", config_path)
        sys.exit(1)
    except yaml.YAMLError as e:
        logger.error(Fore.RED + "❌ ERROR: Gagal membaca file %s: %s", config_path, e)
        sys.exit(1)
    except ValueError as e:
        logger.error(Fore.RED + "❌ ERROR: %s", e)
        sys.exit(1)

def get_location_interactively(loc_config: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    print(Fore.YELLOW + "Pilih lokasi:")
    for key, value in loc_config.items():
        print(f"{key}. {value.get('name', 'Nama Tidak Diketahui')}")

    location_choice = input(Fore.GREEN + "Masukkan nomor: " + Style.RESET_ALL).strip()

    if location_choice not in loc_config:
        raise ValueError(Fore.RED + "Pilihan lokasi tidak valid!")
    return loc_config[location_choice]

def main(args: argparse.Namespace) -> None:
    global LOCATION_CONFIG
    LOCATION_CONFIG = load_config(CONFIG_FILE)

    logger.warning(Fore.CYAN + "\n=== MEMULAI PROSES ===" + Style.RESET_ALL)

    try:
        loc_data: Dict[str, Any]
        if args.location and args.location in LOCATION_CONFIG:
            loc_data = LOCATION_CONFIG[args.location]
            logger.warning(Fore.BLUE + f"Lokasi dipilih dari argumen: {loc_data.get('name', 'N/A')}" + Style.RESET_ALL) # Emoji dihapus
        elif args.location:
            logger.warning(Fore.YELLOW + f"Lokasi '{args.location}' dari argumen tidak ditemukan. Silakan pilih secara interaktif." + Style.RESET_ALL) # Emoji tidak ada di sini
            loc_data = get_location_interactively(LOCATION_CONFIG)
        else:
            loc_data = get_location_interactively(LOCATION_CONFIG)
        
        logger.warning(Fore.BLUE + f"Lokasi yang diproses: {loc_data.get('name', 'N/A')}" + Style.RESET_ALL) # Emoji dihapus

        output_file_name: str
        if args.output:
            output_file_name = args.output
        else:
            output_file_name = input(Fore.GREEN + "Masukkan nama file output (template): " + Style.RESET_ALL).strip()

        if not output_file_name.lower().endswith('.xlsx'):
            output_file_name += ".xlsx"
        logger.warning(f"Nama file output (template) diatur ke: {output_file_name}")

        template_dir = PROJECT_ROOT / TEMPLATE_DIR_NAME
        if not template_dir.exists():
            logger.warning(f"Membuat folder template: {template_dir}")
            template_dir.mkdir(parents=True, exist_ok=True)
        template_file_path = template_dir / output_file_name

        source_file_name: str
        if args.source:
            source_file_name = args.source
        else:
            source_file_name = input(Fore.GREEN + "Masukkan nama file sumber data: " + Style.RESET_ALL).strip()

        if not source_file_name.lower().endswith('.xlsx'):
            source_file_name += ".xlsx"
        
        source_file_path = Path(source_file_name)
        if not source_file_path.is_absolute() and not source_file_path.exists():
             source_file_path = PROJECT_ROOT / source_file_name

        logger.warning(f"File sumber data diatur ke: {source_file_path}")

        use_ref: bool
        if args.use_reference is not None:
            use_ref = args.use_reference
        else:
            use_ref_input = input(Fore.GREEN + "Gunakan referensi (No. Faktur di Customer)? (y/n, default n): " + Style.RESET_ALL).lower().strip()
            use_ref = use_ref_input == 'y'
        logger.warning(f"Gunakan referensi diatur ke: {use_ref}")

        logger.warning(Fore.YELLOW + "\n=== STEP 1: MEMBUAT TEMPLATE ===" + Style.RESET_ALL)
        create_template(str(template_file_path), loc_data)
        if not template_file_path.exists():
            logger.error(Fore.RED + f"Gagal membuat file template di: {template_file_path}")
            raise FileNotFoundError(f"Template file '{template_file_path}' tidak ditemukan setelah proses pembuatan.")

        logger.warning(Fore.YELLOW + "\n=== STEP 2: MEMPROSES DATA CUSTOMER ===" + Style.RESET_ALL)
        process_customer(str(template_file_path), str(source_file_path), use_ref, loc_data.get('id_tku', 'ID_TKU_DEFAULT'))

        logger.warning(Fore.YELLOW + "\n=== STEP 3: MEMPROSES DATA BARANG ===" + Style.RESET_ALL)
        populate_detail_faktur(str(template_file_path), str(source_file_path))

        logger.warning(Fore.YELLOW + "\n=== STEP 4: MELAKUKAN VLOOKUP DATA ===" + Style.RESET_ALL)
        full_vlookup(str(template_file_path), loc_data)

        logger.warning(Fore.GREEN + "\n✅ SEMUA PROSES SELESAI!" + Style.RESET_ALL)
        logger.warning(Fore.BLUE + f"📁 File template yang diproses: {template_file_path}" + Style.RESET_ALL) # Emoji dihapus
        logger.warning(Fore.BLUE + "   Hasil akhir dengan VLOOKUP disimpan di folder 'imp'." + Style.RESET_ALL)

    except FileNotFoundError as e:
        logger.error(Fore.RED + f"\n❌ ERROR FILE TIDAK DITEMUKAN: {e}" + Style.RESET_ALL)
    except ValueError as e:
        logger.error(Fore.RED + f"\n❌ ERROR INPUT TIDAK VALID: {e}" + Style.RESET_ALL)
    except Exception as e:
        logger.exception(Fore.RED + f"\n❌ ERROR UTAMA TIDAK TERDUGA: {e}" + Style.RESET_ALL)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Program untuk memproses data faktur.")
    parser.add_argument(
        "-l", "--location",
        type=str,
        help="Nomor ID lokasi (misal: '1' untuk Surabaya). Jika tidak diberikan, akan ditanyakan."
    )
    parser.add_argument(
        "-s", "--source",
        type=str,
        help="Nama file sumber data Excel (misal: 'data_penjualan.xlsx')."
    )
    parser.add_argument(
        "-o", "--output",
        type=str,
        help="Nama file output untuk template (misal: 'template_faktur.xlsx')."
    )
    parser.add_argument(
        "--use-reference",
        action=argparse.BooleanOptionalAction,
        help="Gunakan kolom referensi (No. Faktur) saat memproses data customer."
    )

    cli_args = parser.parse_args()
    main(cli_args)