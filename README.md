# botmappertamina

Bot otomatisasi input data transaksi MAP (Merchant Apps Pangkalan) Subsidi Tepat Pertamina LPG 3 Kg berbasis Python & Playwright.

---

## Panduan Format File Excel (`template_nik.xlsx`)

File data transaksi menggunakan format Excel (`.xlsx`). File ini **diabaikan oleh git (.gitignore)** demi menjaga privasi dan keamanan data NIK konsumen pangkalan Anda.

Jika file `template_nik.xlsx` belum ada saat bot dijalankan, bot akan **otomatis membuatkan file baru** dengan format header yang benar.

### Struktur Kolom

| Kolom | Nama Kolom | Diisi Oleh | Keterangan |
| :---: | :--- | :---: | :--- |
| **A** | **NIK** | **Pengguna** | Masukkan 16 digit NIK pelanggan KTP. Format cell disarankan **Text** agar angka tidak berubah menjadi format ilmiah (seperti `3.51E+15`). |
| **B** | **STATUS** | *Bot* | Hasil verifikasi: `SUKSES`, `TERDAFTAR`, `BELUM TERDAFTAR`, atau `NIK BERMASALAH`. |
| **C** | **KETERANGAN** | *Bot* | Rincian status: `LUNAS (1 Tabung LPG 3 Kg)`, `Full` (kuota bulanan habis), atau alasan error. |
| **D** | **NAMA PELANGGAN** | *Bot* | Nama resmi konsumen yang diambil otomatis dari server Dukcapil / Pertamina saat transaksi. |
| **E** | **TERAKHIR DIGUNAKAN** | *Bot* | Tanggal dan jam transaksi berhasil (format: `YYYY-MM-DD HH:MM:SS`). |

### Contoh Isi File Excel:

| NIK | STATUS | KETERANGAN | NAMA PELANGGAN | TERAKHIR DIGUNAKAN |
| :--- | :--- | :--- | :--- | :--- |
| 3510212206460003 | SUKSES | LUNAS (1 Tabung LPG 3 Kg) | SOEPENO ROMANI | 2026-10-30 08:30:15 |
| 3510217006790038 | TERDAFTAR | Full | SIYAMAH | 2026-10-21 06:54:04 |
| 3510155208790000 | BELUM TERDAFTAR | Pelanggan Tidak Terdaftar | | |
| 3510211906880003 | | | | |

---

## Mekanisme Kerja Bot

1. **Prioritas Antrean Cerdas (*Smart Queue Sorting*)**:
   - NIK yang kolom `TERAKHIR DIGUNAKAN`-nya masih kosong (belum pernah dipakai) akan diproses **paling awal**.
   - NIK yang sudah pernah dipakai akan diurutkan dari **tanggal yang paling lampau / terlama**.
2. **Penyelesaian Captcha Slider Otomatis**:
   - Membaca posisi puzzle dan memicu verifikasi secara instan tanpa perlu digeser manual.
3. **Pilihan Kategori Konsumen**:
   - Otomatis memilih opsi **"Rumah Tangga"** jika konsumen terdaftar di lebih dari satu jenis pelanggan (Usaha Mikro / Pengecer / Rumah Tangga).
4. **Auto-Lengkapi Data Pelanggan**:
   - Jika muncul peringatan data Dukcapil belum lengkap, bot otomatis memecah NIK menjadi Tempat Lahir (kabupaten/kota via database wilayah) dan Tanggal Lahir untuk mengisi form secara otomatis.
5. **Anti-Throttling & Jeda Realistis**:
   - Mengetik NIK dengan kecepatan ketik alami dan jeda aman antar transaksi agar tidak terkena limit rate Pertamina.

---

## Cara Menjalankan Bot

1. Buka dan isi NIK pelanggan pada kolom **A** di file `template_nik.xlsx`.
2. **PENTING**: **Tutup aplikasi Microsoft Excel / WPS** sebelum menjalankan bot. (Windows akan mengunci file jika Excel masih terbuka).
3. Klik 2x file **`JALANKAN_BOT.bat`**.
4. Bot akan otomatis membuka browser, login dengan kredensial dari `config.json`, dan memproses antrean NIK satu per satu.
5. Progres transaksi akan langsung tersimpan secara aman (*real-time*) ke file Excel.

---

## Setup di Laptop Baru (Hanya 1x)

1. Pasang Python (versi 3.10 atau lebih baru) dari [python.org](https://www.python.org/downloads/).
   - **Wajib centang kotak "Add python.exe to PATH"** pada installer Python.
2. Salin folder project ini ke laptop baru.
3. Klik 2x file **`INSTALL_DI_LAPTOP_BARU.bat`** untuk memasang seluruh dependensi (`playwright`, `openpyxl`, browser Chromium).
4. Selesai! Anda bisa langsung menjalankan bot melalui `JALANKAN_BOT.bat`.
