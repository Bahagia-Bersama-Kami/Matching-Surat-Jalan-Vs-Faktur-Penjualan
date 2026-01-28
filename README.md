# Matching-Surat-Jalan-Vs-Faktur-Penjualan
Accurate Transaction Reconciler (Pengiriman Pesanan vs Faktur Penjualan)
Alat otomatisasi berbasis Python untuk mengekstrak, menggabungkan, dan menganalisis data laporan dari software akuntansi Accurate. Repositori ini dirancang khusus untuk menangani file .xls "palsu" (format XML) — Simpan -> Excel Table (XML) — yang dihasilkan oleh Accurate dan melakukan rekonsiliasi otomatis antara transaksi Pengiriman Pesanan (Delivery Order) dan Faktur Penjualan (Sales Invoice).

# Fitur Utama
1. Otomatisasi Alur Kerja: Cukup jalankan satu skrip untuk memproses puluhan file sekaligus.
2. Accurate XML Parser: Mengonversi file .xls hasil ekspor Accurate (yang sebenarnya adalah XML) menjadi DataFrame Pandas yang bersih.
3. Analisis Kecocokan (Matching Logic):
- Mendeteksi pasangan transaksi berdasarkan Nomor Faktur.
- Identifikasi selisih angka (toleransi > 5).
- Identifikasi perbedaan tanggal antara pengiriman dan faktur.
- Deteksi transaksi yang tidak memiliki pasangan (pasangan hilang).
4. Laporan Excel Profesional:
- Output yang rapi dengan pewarnaan otomatis (Hijau untuk ringkasan, Merah untuk error/selisih, Biru untuk data asli).
- Fitur Auto-fit column width untuk kenyamanan pembacaan.
- Format angka desimal standar Indonesia (menggunakan titik untuk ribuan dan koma untuk desimal).

# Persyaratan Sistem
Pastikan Anda telah menginstal Python 3.x dan pustaka yang diperlukan:
```bash
pip install pandas openpyxl
```

# Cara Penggunaan
1. Siapkan Data: Masukkan semua file laporan .xls hasil ekspor dari Accurate ke dalam folder Input/.
2. Jalankan Program: Klik dua kali atau jalankan skrip utama melalui terminal:
```bash
python "Jalankan Analisis.py"
```

# Skrip Tambahan (Add-ons)
Dalam folder Dapur/, tersedia alat bantu lainnya:
1. Merger: Menggabungkan banyak file Excel yang memiliki header yang sama menjadi satu file.
2. Sheet Extractor: Memecah satu file Excel yang memiliki banyak sheet menjadi banyak file Excel terpisah (satu file per sheet).
