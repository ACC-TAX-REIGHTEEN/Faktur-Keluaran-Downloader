# Sistem Ekstraksi Data Invoice JSON ke Excel

Repositori ini berisi perangkat alat untuk mengekstraksi data daftar invoice dari API berbasis web (dalam format JSON) dan mengonversinya menjadi laporan Excel yang terstruktur dan mudah dibaca.

## Komponen Utama

Sistem ini terdiri dari dua komponen utama yang bekerja secara berurutan:

1. **Esktrak Script Json to TXT.js**: Skrip JavaScript (Bookmarklet/Console) untuk menangkap respons API dari browser dan menyimpannya sebagai file teks.
2. **Ekstrak Json TXT to XLSX.py**: Skrip Python untuk mengolah file teks JSON tersebut menjadi file Excel (.xlsx) dengan pemetaan kolom dalam bahasa Indonesia.

## Alur Kerja dan Penggunaan

### Langkah 1: Pengambilan Data (JavaScript)

Skrip JavaScript berfungsi untuk mengintersepsi permintaan jaringan pada browser.

1. Buka halaman web yang memuat daftar invoice.
2. Buka Developer Tools (F12) dan masuk ke tab Console, atau simpan skrip sebagai Bookmarklet.
3. Jalankan skrip tersebut.
4. Masukkan jumlah baris yang ingin ditampilkan pada prompt yang muncul (misalnya: 1000).
5. Lakukan refresh atau trigger pemuatan data pada web tersebut.
6. Browser akan secara otomatis mengunduh file bernama `outputinvoice_list.txt`.

### Langkah 2: Konversi ke Excel (Python)

Skrip Python akan membaca file hasil unduhan dan melakukan pembersihan data.

1. Pastikan file `outputinvoice_list.txt` berada di direktori yang sama dengan skrip Python.
2. Jalankan skrip Python:
   ```bash
   python "Ekstrak Json TXT to XLSX.py"
