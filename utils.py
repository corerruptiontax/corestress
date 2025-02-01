# utils.py
from colorama import Fore, Style
from datetime import datetime
import pandas as pd

def convert_date(tgl_input):
    """
    Mengkonversi berbagai format tanggal ke DD/MM/YYYY
    
    Parameter:
    tgl_input (str/pd.Timestamp): Input tanggal
    
    Return:
    str: Tanggal dalam format DD/MM/YYYY atau 'INVALID_DATE'
    """
    try:
        # Handle datetime object dari Excel
        if isinstance(tgl_input, pd.Timestamp):
            return tgl_input.strftime("%d/%m/%Y")
        
        tgl_str = str(tgl_input).strip()
        
        # Handle format YYYY-MM-DD
        if "-" in tgl_str and len(tgl_str.split("-")[0]) == 4:
            return datetime.strptime(tgl_str.split()[0], "%Y-%m-%d").strftime("%d/%m/%Y")
        
        # Handle format DD MMM YYYY
        bulan_translation = {
            'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr',
            'Mei': 'May', 'Jun': 'Jun', 'Jul': 'Jul', 'Agu': 'Aug',
            'Sep': 'Sep', 'Okt': 'Oct', 'Nov': 'Nov', 'Des': 'Dec'
        }
        
        # Normalisasi input
        normalized = tgl_str.replace('-', ' ').replace('/', ' ')
        for id_bulan, en_bulan in bulan_translation.items():
            normalized = normalized.replace(id_bulan, en_bulan)
            
        # Coba parse
        return datetime.strptime(normalized, "%d %b %Y").strftime("%d/%m/%Y")
    
    except Exception as e:
        print(Fore.YELLOW + f"⚠️ Gagal konversi: {tgl_input} ({str(e)})" + Style.RESET_ALL)
        return "INVALID_DATE"
