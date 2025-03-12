# CoreStress  
 
Repositori ini berisi kumpulan skrip Python yang dirancang untuk meng-otomatis-kan seluruh proses pengolahan data dan interaksi dengan file Excel. Proyek ini bertujuan untuk meningkatkan efisiensi dan mengurangi kesalahan manual dalam pengelolaan data.

## Deskripsi

Proyek ini merupakan sistem otomatisasi yang mengintegrasikan beberapa modul untuk memproses data secara efisien. Skrip-skrip ini bekerja sama untuk melakukan berbagai tugas, termasuk:

1. Interaksi dengan Basis Data: Mengelola dan mengambil data dari basis data menggunakan db.py.
2. Pengelolaan Barang: Mengelola informasi terkait barang melalui goods.py.
3. Pengelolaan Pelanggan: Memproses data pelanggan dengan cust.py.
4. Pengaturan dan Konfigurasi: Menyediakan template dan pengaturan awal melalui setupcore.py.
5. Pengolahan Data Excel: Menggunakan openpyxl dan pandas untuk membaca, menulis, dan memanipulasi file Excel.
6. Dengan menggunakan skrip ini, pengguna dapat dengan mudah mengotomatiskan tugas-tugas yang berulang dan fokus pada analisis data yang lebih mendalam.
  
## Instalasi  
Ikuti langkah-langkah berikut untuk menginstal Python dan menjalankan skrip:  
  
1. **Instalasi Python**  
   - Unduh dan instal Python dari [python.org](https://www.python.org/downloads/).  
   - Pastikan untuk mencentang opsi "Add Python to PATH" saat instalasi.  
  
2. **Cek Versi Python**  
   - Buka Command Prompt dan jalankan:  
     ```python --version```
     
3. **Instalasi pip**  
   - Pip biasanya sudah terinstal dengan Python. Untuk memeriksa, jalankan:  
     ```pip --version```
     
4. **Instalasi Dependensi**  
   - Instal dependensi yang diperlukan dengan menjalankan:  
     ```pip install -r requirements.txt```

5. **Menjalankan Skrip**  
   - Clone repositori ke mesin lokal Anda:
     
     ```git clone https://github.com/ssyahbandi/Corestress```

   - Masuk ke direktori proyek:
     
     ```cd Corestress```

   - Instal dependensi yang diperlukan:
     
     ```pip install -r requirements.txt```

   - Jalankan skrip utama:

      ```python main.py```

### Formula Excel  
- Di sel A2, gunakan formula berikut: 

  ```=IF(COUNTIFS(A$2:A2, A2, B$2:B2, B2, D$2:D2, D2, C$2:C2,C2 )=1, MAX(J$1:J1)+1, J1)```

- Saya Menyebutnya ini ```source_file.xlsx```
    ![image](https://github.com/user-attachments/assets/28b99137-0f28-42b8-b419-6be4bc55737e)

# Change Logs

---

📌 Corestress v1.1 - 12/03/2025
🔄 Perubahan:
- ✨ **Perubahan pada Header di Template Excel**
- 🛠️ **Perubahan pada VLOOKUP Data**
- 📂 **Perubahan dalam Penyimpanan File**

---

📌 Corestress v1.0 - 11/02/2025
🔄 Fitur Awal:
- ✅ **Cek eksistensi file sebelum diproses**
- 📝 **Docstring di semua fungsi utama**
- 📅 **Format cell tanggal otomatis (DD/MM/YYYY)**
- 🔍 **Kode Barang yang mengandung "text" tertentu**
- 📏 **Satuan Barang yang mengandung "text" tertentu**
- 📂 **File dan folder terstruktur**
  
## Kontribusi  
Jika Anda ingin berkontribusi pada proyek ini, silakan fork repositori ini dan buat pull request.  
  
## Lisensi  
Proyek ini dilisensikan di bawah [MIT License](LICENSE).  
