import pandas as pd
from .logger import log_info

def normalize_format(df):
    # Kolom standar yang diharapkan oleh data_processor
    standard_columns = ['No. Pelanggan', 'Nama Pelanggan', 'No. Faktur', 'Tgl. Faktur', 'Nama Barang', 'Qty', 'DPP+PPN']

    # Deteksi header asli (DPP atau DPP+PPN)
    is_dpp_header = 'DPP' in df.columns and 'DPP+PPN' not in df.columns

    # Deteksi format berdasarkan kolom yang ada
    if 'Cabang' in df.columns:
        log_info("Format terdeteksi: Format 2 (dengan kolom Cabang, Kategori Pelanggan, Kota, DPP)")
        # Format 2: DPP akan digunakan untuk perhitungan yang berbeda
        if is_dpp_header:
            df = df.rename(columns={'DPP': 'DPP+PPN'})  # Ganti nama untuk konsistensi, tapi kita tahu ini DPP
        # Hapus baris Total
        df = df[~df['No. Faktur'].str.contains('Total', na=False)]
        # Hapus kolom tambahan
        df = df.drop(columns=['Cabang', 'Kategori Pelanggan', 'Kota'], errors='ignore')
    elif any(df['No. Faktur'].str.contains('Total', na=False)):
        log_info("Format terdeteksi: Format 3 (dengan baris Total)")
        # Format 3: Hapus baris Total
        df = df[~df['No. Faktur'].str.contains('Total', na=False)]
    else:
        log_info("Format terdeteksi: Format 1 (tanpa baris Total)")

    # Pastikan semua kolom standar ada
    missing_columns = [col for col in standard_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Kolom yang diharapkan tidak ditemukan setelah normalisasi: {missing_columns}. Kolom yang ada: {list(df.columns)}")

    # Kembalikan hanya kolom yang diperlukan dan informasi header asli
    return df[standard_columns], is_dpp_header