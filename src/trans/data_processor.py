import pandas as pd
from .logger import log_info, log_warning, log_important

def process_data(df, is_dpp_header=False):
    current_customer_number = None
    current_customer_name = None
    current_invoice_number = None
    previous_invoice_number = None
    result_data = []
    deleted_rows = []
    row_count = 0
    auto_delete_minus = False
    auto_delete_zero = False
    deleted_zero = 0
    deleted_minus = 0
    minus_item_name = None
    zero_item_name = None
    indices_to_delete = []
    is_first_row_of_invoice = True
    invoice_dates = {}
    invoice_rows = {}

    # Tahap 1: Simpan semua Tgl. Faktur dari file raw berdasarkan No. Faktur
    current_invoice = None
    for index, row in df.iterrows():
        invoice_number = str(row['No. Faktur']).strip()
        if invoice_number != '':
            current_invoice = invoice_number
        tgl_faktur = str(row['Tgl. Faktur']).strip()
        if tgl_faktur != '' and current_invoice:
            try:
                tgl_faktur = pd.to_datetime(tgl_faktur).strftime('%d %b %Y')
                invoice_dates[current_invoice] = tgl_faktur
            except (ValueError, TypeError) as e:
                log_info(f"Error saat mengubah format tanggal di file raw - {e}, menggunakan tanggal asli: {tgl_faktur}")

    # Tahap 2: Proses semua baris dan tandai baris yang akan dihapus
    for index, row in df.iterrows():
        row_count += 1
        log_info(f"Memproses baris {row_count}: No. Pelanggan='{row['No. Pelanggan']}', No. Faktur='{row['No. Faktur']}', Nama Barang='{row['Nama Barang']}', Qty='{row['Qty']}', DPP+PPN='{row['DPP+PPN']}'")

        # Cek apakah baris benar-benar kosong (semua kolom kosong)
        is_empty_row = all(str(row[col]).strip() == '' for col in df.columns)

        if is_empty_row:
            log_info(f"Baris {row_count} kosong, menambahkan baris kosong ke output")
            result_data.append([None] * 10)
            deleted_rows.append([None] * 10)
            is_first_row_of_invoice = True
            continue

        # Update No. Pelanggan dan Nama Pelanggan jika ada
        if str(row['No. Pelanggan']).strip() != '':
            current_customer_number = row['No. Pelanggan']
            current_customer_name = row['Nama Pelanggan']
            log_info(f"Baris {row_count}: Update pelanggan - {current_customer_number}, {current_customer_name}")

        # Pastikan data pelanggan ada
        if not current_customer_number or not current_customer_name:
            log_important(f"Baris {row_count}: Tidak ada data pelanggan, melewati baris")
            continue

        # Ambil No. Faktur
        invoice_number = str(row['No. Faktur']).strip()
        if invoice_number != '':
            current_invoice_number = invoice_number
        else:
            invoice_number = current_invoice_number if current_invoice_number else ''
            log_info(f"Baris {row_count}: No. Faktur kosong, menggunakan No. Faktur sebelumnya: {invoice_number}")

        if invoice_number == '':
            log_important(f"Baris {row_count}: No. Faktur masih kosong setelah pengisian, melewati baris")
            continue

        if invoice_number not in invoice_rows:
            invoice_rows[invoice_number] = []
        invoice_rows[invoice_number].append(len(result_data))

        nama_barang = str(row['Nama Barang']).strip()
        if nama_barang == '':
            log_important(f"Baris {row_count}: Nama Barang kosong, melewati baris")
            continue

        if invoice_number != previous_invoice_number:
            previous_invoice_number = invoice_number
            is_first_row_of_invoice = True
            log_info(f"Baris {row_count}: No. Faktur berubah ke {invoice_number}")
        else:
            is_first_row_of_invoice = False

        try:
            dpp_ppn = float(str(row['DPP+PPN']).replace(',', ''))
            qty = float(row['Qty'])
        except (ValueError, TypeError) as e:
            log_important(f"Baris {row_count}: Error pada DPP+PPN atau Qty - {e}, melewati baris")
            continue

        # Hitung Harga DPP berdasarkan header asli
        if is_dpp_header:
            harga_dpp = dpp_ppn / qty if qty != 0 else 0
        else:
            harga_dpp = dpp_ppn / qty / 1.11 if qty != 0 else 0

        total_dpp = harga_dpp * qty
        ppn = total_dpp * 0.11

        tgl_faktur = invoice_dates.get(invoice_number, '')
        if not is_first_row_of_invoice:
            tgl_faktur = ''

        row_to_add = [
            current_customer_number,
            current_customer_name,
            invoice_number,
            tgl_faktur,
            nama_barang,
            harga_dpp,
            qty,
            total_dpp,
            ppn,
            0
        ]
        result_data.append(row_to_add)
        log_info(f"Baris {row_count}: Data ditambahkan - {nama_barang}")

        if harga_dpp < 0 or total_dpp < 0:
            if auto_delete_minus and minus_item_name and minus_item_name.lower() in nama_barang.lower():
                log_important(f"Baris {row_count} dihapus otomatis karena nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} dihapus otomatis karena nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ❌")
                deleted_rows.append(row_to_add)
                indices_to_delete.append(len(result_data) - 1)
                deleted_minus += 1
                continue
            log_warning(f"PERINGATAN: Baris {row_count} memiliki nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) - Kemungkinan retur di Nama Barang '{nama_barang}'")
            log_info(f"PERINGATAN: Baris {row_count} memiliki nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) - Kemungkinan retur di Nama Barang '{nama_barang}' ⚠️")
            choice = input(f"Hapus baris {row_count} karena nilai minus? (y/n/y-all): ").strip().lower()
            if choice == 'y-all':
                auto_delete_minus = True
                minus_item_name = nama_barang
                log_important(f"Baris {row_count} dihapus karena nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} dihapus karena nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ❌")
                deleted_rows.append(row_to_add)
                indices_to_delete.append(len(result_data) - 1)
                deleted_minus += 1
                continue
            elif choice == 'y':
                log_important(f"Baris {row_count} dihapus karena nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} dihapus karena nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ❌")
                deleted_rows.append(row_to_add)
                indices_to_delete.append(len(result_data) - 1)
                deleted_minus += 1
                continue
            else:
                log_important(f"Baris {row_count} tetap disertakan meskipun nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} tetap disertakan meskipun nilai minus (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ✅")

        if harga_dpp == 0 or total_dpp == 0:
            if auto_delete_zero and zero_item_name and zero_item_name.lower() in nama_barang.lower():
                log_important(f"Baris {row_count} dihapus otomatis karena nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} dihapus otomatis karena nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ❌")
                deleted_rows.append(row_to_add)
                indices_to_delete.append(len(result_data) - 1)
                deleted_zero += 1
                continue
            log_warning(f"PERINGATAN: Baris {row_count} memiliki nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) - Kemungkinan bonus di Nama Barang '{nama_barang}'")
            log_info(f"PERINGATAN: Baris {row_count} memiliki nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) - Kemungkinan bonus di Nama Barang '{nama_barang}' ⚠️")
            choice = input(f"Hapus baris {row_count} karena nilai 0? (y/n/y-all): ").strip().lower()
            if choice == 'y-all':
                auto_delete_zero = True
                zero_item_name = nama_barang
                log_important(f"Baris {row_count} dihapus karena nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} dihapus karena nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ❌")
                deleted_rows.append(row_to_add)
                indices_to_delete.append(len(result_data) - 1)
                deleted_zero += 1
                continue
            elif choice == 'y':
                log_important(f"Baris {row_count} dihapus karena nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} dihapus karena nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ❌")
                deleted_rows.append(row_to_add)
                indices_to_delete.append(len(result_data) - 1)
                deleted_zero += 1
                continue
            else:
                log_important(f"Baris {row_count} tetap disertakan meskipun nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp})")
                log_info(f"Baris {row_count} tetap disertakan meskipun nilai 0 (Harga DPP: {harga_dpp}, Total DPP: {total_dpp}) ✅")

    # Tahap 3: Hapus baris yang ditandai dari result_data
    indices_to_delete = sorted(indices_to_delete, reverse=True)
    for idx in indices_to_delete:
        if 0 <= idx < len(result_data):
            del result_data[idx]

    # Tahap 4: Tambahkan pemisah transaksi dan buat indeks invoice
    intermediate_data = result_data.copy()
    result_data = []
    current_invoice = None
    current_customer = None
    invoice_customer_pairs = []  # Untuk menyimpan pasangan (invoice, customer)
    invoice_indices = {}  # Untuk melacak indeks awal setiap invoice
    idx = 0

    # Tahap 4.1: Identifikasi pasangan invoice dan pelanggan yang tersisa dan simpan urutannya
    for row in intermediate_data:
        if row[0] is None:  # Baris kosong
            continue
        invoice_number = row[2]  # Kolom No. Faktur
        customer_name = row[1]  # Kolom Nama Pelanggan
        if invoice_number != current_invoice or customer_name != current_customer:
            current_invoice = invoice_number
            current_customer = customer_name
            invoice_customer_pairs.append((invoice_number, customer_name))

    # Tidak mengurutkan invoice_customer_pairs, sehingga mengikuti urutan kemunculan di file raw

    # Tahap 4.2: Tambahkan data dan pemisah transaksi
    current_invoice = None
    current_customer = None
    for row in intermediate_data:
        if row[0] is None:  # Baris kosong (pemisah pelanggan)
            result_data.append(row)
            idx += 1
            continue

        invoice_number = row[2]  # Kolom No. Faktur
        customer_name = row[1]  # Kolom Nama Pelanggan
        if invoice_number != current_invoice or customer_name != current_customer:
            # Tambahkan pemisah transaksi hanya jika invoice atau pelanggan sebelumnya ada
            if current_invoice is not None or current_customer is not None:
                result_data.append([None] * 10)
                idx += 1
            current_invoice = invoice_number
            current_customer = customer_name
            invoice_indices[(current_invoice, current_customer)] = idx

        result_data.append(row)
        idx += 1

    # Tahap 5: Hitung ulang nomor baris berdasarkan pasangan No. Faktur dan Nama Pelanggan yang tersisa
    final_result_data = []
    current_invoice = None
    current_customer = None
    for idx, row in enumerate(result_data):
        if row[0] is None:  # Baris kosong
            final_result_data.append(row)
            continue
        invoice_number = row[2]  # Kolom No. Faktur
        customer_name = row[1]  # Kolom Nama Pelanggan
        if invoice_number != current_invoice or customer_name != current_customer:
            current_invoice = invoice_number
            current_customer = customer_name
        new_row = list(row)
        # Gunakan invoice_customer_pairs untuk menentukan nomor baris yang berurutan
        new_row[9] = invoice_customer_pairs.index((current_invoice, current_customer)) + 1  # Kolom Baris
        final_result_data.append(new_row)

    # Tahap 6: Pastikan Tgl. Faktur ada di baris pertama setiap kelompok No. Faktur
    current_invoice = None
    for idx, row in enumerate(final_result_data):
        if row[0] is None:  # Baris kosong
            continue
        invoice_number = row[2]  # Kolom No. Faktur
        if invoice_number != current_invoice:
            current_invoice = invoice_number
            if invoice_number in invoice_dates:
                final_result_data[idx][3] = invoice_dates[invoice_number]
        else:
            final_result_data[idx][3] = ''

    # Tahap 7: Catat jika sebuah invoice hanya berisi baris yang dihapus
    for invoice, indices in invoice_rows.items():
        remaining_rows = [i for i in indices if i not in indices_to_delete]
        if not remaining_rows:
            log_info(f"Invoice {invoice} hanya berisi baris yang dihapus (Bonus/Retur) dan tidak muncul di hasil akhir")

    # Tahap 8: Hitung total keseluruhan DPP dan PPN
    total_dpp = 0
    total_ppn = 0
    for row in final_result_data:
        if row[0] is None:  # Lewati baris kosong
            continue
        total_dpp += row[7]  # Kolom Total DPP
        total_ppn += row[8]  # Kolom PPN

    return final_result_data, deleted_zero, deleted_minus, deleted_rows, total_dpp, total_ppn