# botmappertamina

Bot otomatisasi input data transaksi MAP (Merchant Apps Pangkalan) Subsidi Tepat Pertamina LPG 3 Kg berbasis Python & Playwright.

## Fitur Utama

- **Login Otomatis & Manajemen Sesi**: Login otomatis menggunakan kredensial pangkalan dan menyimpan profil browser agar reCAPTCHA tidak memblokir.
- **Bypass Captcha Slider / Puzzle**: Mengatasi puzzle captcha keamanan penjualan secara instan & presisi tanpa geser manual.
- **Sorting Otomatis Berdasarkan Terlama**: Mengurutkan antrean input dari data yang belum pernah/terlama digunakan.
- **Auto-Resolve Multi-Kategori**: Otomatis memilih "Rumah Tangga" jika pelanggan memiliki beberapa kategori (Usaha Mikro / Pengecer / Rumah Tangga).
- **Auto-Update Data Pelanggan**: Mengurai NIK secara otomatis (kode wilayah & tanggal/bulan/tahun lahir) untuk mengisi form "Lengkapi Data Pelanggan" jika data Dukcapil belum lengkap.
- **Auto-Handle Peringatan NIB & Kuota Limit**: Menangani peringatan NIB Usaha Mikro dan menandai status "Full" jika kuota bulanan pelanggan sudah habis.
- **Pencatatan Real-time ke Excel**: Menyimpan Nama Pelanggan, Status, Keterangan, dan Timestamp transaksi secara aman (anti-crash / auto-retry).

## Cara Penggunaan

1. Buka dan masukkan daftar NIK ke kolom A pada file `template_nik.xlsx`.
2. Pastikan file Excel ditutup sebelum menjalankan bot.
3. Klik 2x file `JALANKAN_BOT.bat` untuk memulai proses input transaksi.

## Setup di Laptop Baru

1. Pastikan Python sudah terpasang (centang opsi **Add python.exe to PATH** saat instalasi).
2. Klik 2x file `INSTALL_DI_LAPTOP_BARU.bat` untuk memasang seluruh dependensi secara otomatis.
3. Jalankan bot melalui `JALANKAN_BOT.bat`.
