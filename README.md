# CoreStress  
  
[Saya sedang malas merapihkan, nanti saya update, cara pakainya]  
  
## Deskripsi  
Repositori ini berisi skrip Python untuk memproses data dari file Excel. Skrip ini dirancang untuk membantu dalam pengolahan data faktur dan barang.  
  
## Format File Sumber  
Sebelum menjalankan `move.py`, pastikan format file sumber Anda adalah sebagai berikut:  
  
**Contoh Format:**  
| No. Pelanggan | Nama Pelanggan | No. Faktur | Tgl. Faktur | ... | Barang |  
|----------------|----------------|-------------|--------------|-----|--------|  
| P-PT-LIN-ORE   | [Nama Pelanggan] | CV412178   | 31/12/2024   | ... | [Nama barang] |  
| P-IB-SUR-GRE   | [Nama Pelanggan] | CV412306   | 24/12/2024   | ... | [Nama barang] |  
| ...            | ...            | ...         | ...          | ... | ...    |  
  
### Formula Excel  
Di sel A1, gunakan formula berikut:  
"=IF(COUNTIFS(B$2:B, B2, C$2:C, E2, E$2:E, E2, D$2:D, D2)=1, MAX(A$1:A1)+1, A1)"

## Menjalankan Skrip  
Untuk menjalankan skrip, gunakan perintah berikut di terminal:  

## Skrip yang Tersedia  
- `cust.py`: Skrip untuk memproses data faktur.  
- `goods.py`: Skrip untuk memproses data barang.  
- `setup.py`: Skrip untuk menyiapkan lingkungan kerja.  
- `requirements.txt`: Daftar dependensi yang diperlukan.  
- `FF.txt`: Contoh hasil output dari `setup.py` yang digunakan saat menjalankan `cust.py` dan `goods.py`.  
- `Faktur Pajak Desember 2024.xlsx`: Contoh file sumber yang digunakan.  
  
## Instalasi  
Ikuti langkah-langkah berikut untuk menginstal Python dan menjalankan skrip:  
  
1. **Instalasi Python**  
   - Unduh dan instal Python dari [python.org](https://www.python.org/downloads/).  
   - Pastikan untuk mencentang opsi "Add Python to PATH" saat instalasi.  
  
2. **Cek Versi Python**  
   - Buka Command Prompt dan jalankan:  
     ```python --version```
  bash
pip --version
