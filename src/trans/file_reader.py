import pandas as pd

def read_raw_file(file_name, sheet_name=None):
    # Jika sheet_name adalah None, baca sheet pertama secara eksplisit
    if sheet_name is None:
        # Baca file Excel dan ambil sheet pertama
        excel_file = pd.ExcelFile(f"{file_name}.xlsx")
        sheet_name = excel_file.sheet_names[0]  # Ambil nama sheet pertama
    return pd.read_excel(f"{file_name}.xlsx", sheet_name=sheet_name, dtype=str, keep_default_na=False)