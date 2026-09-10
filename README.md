# Bot Otomasi Perbaikan Data Sensus Ekonomi 2026 (Fasih BPS)

Bot otomasi berbasis Python dan Selenium untuk melakukan perbaikan data survei secara massal pada aplikasi **Fasih BPS (Sensus Ekonomi 2026)**.

---

## 📌 Fitur Utama

- **Otomatisasi Penuh Formulir**:
  - Membuka URL assignment langsung dalam mode edit.
  - Mengubah status keberadaan pada **SE2026 - P** menjadi *"1. Ditemukan"*.
  - Mengakses **SE2026 - L BLOK II** $\rightarrow$ membuka *"Isi/Lihat Rincian"* $\rightarrow$ mengubah status keberadaan usaha menjadi *"1. Ditemukan"*.
  - Mengirimkan formulir melalui alur konfirmasi **Kirim $\rightarrow$ Submit Paksa $\rightarrow$ Konfirmasi**.
- **Sistem Checkpoint & Resumable**:
  - Progres tersimpan secara *real-time* dalam file JSON checkpoint.
  - Jika bot dihentikan sewaktu-waktu (`Ctrl + C`), bot akan melanjutkan dari data terakhir tanpa mengulang baris yang sudah sukses.
- **Auto Re-Login 2 Tahap (SSO BPS & Keycloak)**:
  - Otomatis mendeteksi token/sesi kedaluwarsa.
  - Melakukan klik otomatis pada SSO Fasih dan login Keycloak BPS tanpa campur tangan pengguna.
- **Auto-Recovery Rate Limit & Bot Detection**:
  - Mengelola *Rate Limit Exceeded* dengan *exponential backoff*.
  - Menangani proteksi *Bot Detected* (cooldown otomatis 5 menit dan klik *Kembali*).
- **Mode Khusus Retry**:
  - Menyediakan opsi `--retry-failed` untuk menyapu bersih data yang sempat gagal tanpa perlu memindai ulang ribuan data sukses.
- **Pembersihan Screenshot Otomatis**:
  - Mengambil tangkapan layar jika terjadi error untuk investigasi.
  - Otomatis menghapus screenshot error lama ketika data tersebut berhasil diperbaiki (*self-healing*).

---

## 📋 Prasyarat Sistem

Sebelum menjalankan proyek ini, pastikan sistem Anda memenuhi kebutuhan berikut:

1. **Sistem Operasi**: Windows 10/11.
2. **Python**: Versi `3.10` atau lebih baru.
3. **Google Chrome**: Browser resmi terinstal di sistem.
4. **Koneksi VPN BPS**: Aktif dan terhubung ke jaringan internal BPS (agar domain `fasih-sm.bps.go.id` dapat diakses).
5. **Git**: Untuk melakukan clone repository.

---

## 🚀 Panduan Instalasi (Clone & Setup)

### 1. Clone Repository
Buka terminal (PowerShell atau Command Prompt), lalu jalankan:
```bash
git clone https://github.com/RahmadFahrurrozi/automated-fasih-sm.git
cd automated-fasih-sm
```

### 2. Buat Virtual Environment
Buat lingkungan virtual Python agar dependensi terisolasi:
```bash
python -m venv .venv
```

### 3. Aktivasi Virtual Environment
- **PowerShell**:
  ```powershell
  .venv\Scripts\Activate.ps1
  ```
  *(Jika muncul error Execution Policy, jalankan `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` terlebih dahulu)*
- **Command Prompt (CMD)**:
  ```cmd
  .venv\Scripts\activate.bat
  ```

### 4. Install Dependensi
Pasang seluruh pustaka yang dibutuhkan:
```bash
pip install -r requirements.txt
```

---

## ⚙️ Konfigurasi & Persiapan Browser

Bot ini berjalan menggunakan metode **Chrome Remote Debugging**. Hal ini memungkinkan bot mengontrol browser dengan profil dan sesi login pengguna yang sudah ada tanpa konflik sesi.

### 1. Jalankan Chrome dengan Port Debugging
Jalankan file batch:
```cmd
start_chrome_ozifazh.bat
```
Atau jalankan manual via terminal:
```cmd
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\chrome_profile"
```

### 2. Login ke Fasih BPS
1. Pastikan **VPN BPS** sudah aktif.
2. Pada jendela Chrome yang baru terbuka tersebut, buka [fasih-sm.bps.go.id](https://fasih-sm.bps.go.id).
3. Lakukan login akun BPS Anda hingga berhasil masuk ke dashboard survei.
4. **Biarkan jendela Chrome tersebut tetap terbuka.**

---

## 📊 Menyiapkan Data Excel

1. Masukkan file data Excel ke dalam folder `data/` (contoh: `data/030-fix-rozi.xlsx`).
2. Pastikan file Excel memiliki kolom target URL assignment:
   - Nama kolom: **`assignment_id_usaha_bku`**

---

## 🏃 Cara Menjalankan Bot

Tersedia skrip batch praktis satu klik untuk berbagai kebutuhan:

### 1. Eksekusi Normal (Memproses Seluruh Data)
Memproses seluruh baris data dari awal hingga akhir, otomatis melewati data yang sudah tercatat sukses di checkpoint:
```cmd
run_bot.bat
```
Atau via python:
```bash
python bot_fasih_perbaikan.py
```

### 2. Eksekusi Khusus Data Gagal / Retry
Hanya memfilter dan mengeksekusi ulang baris-baris data yang berstatus gagal atau belum selesai:
```cmd
retry_errors.bat
```
Atau via python:
```bash
python bot_fasih_perbaikan.py --retry-failed
```

### 3. Uji Coba Single URL
Untuk menguji coba alur perbaikan pada 1 URL spesifik sebelum menjalankan data massal:
```cmd
test_url.bat
```
Atau via python:
```bash
python bot_fasih_perbaikan.py --url "https://fasih-sm.bps.go.id/app/assignment/.../edit"
```

---

## 🛠️ Opsi Argumen Perintah (CLI)

Anda dapat menambahkan argumen berikut saat menjalankan `python bot_fasih_perbaikan.py`:

| Argumen | Penjelasan | Contoh Penggunaan |
|---|---|---|
| `--excel` | Menentukan path file data Excel tertentu | `python bot_fasih_perbaikan.py --excel "data/nama_file.xlsx"` |
| `--test` | Hanya menguji coba 1 baris pertama | `python bot_fasih_perbaikan.py --test` |
| `--limit <N>` | Membatasi jumlah baris yang diproses | `python bot_fasih_perbaikan.py --limit 50` |
| `--start <N>` | Memulai proses dari indeks baris tertentu (0-indexed) | `python bot_fasih_perbaikan.py --start 100` |
| `--retry-failed` | Hanya memproses data yang gagal atau belum sukses | `python bot_fasih_perbaikan.py --retry-failed` |
| `--reset-checkpoint` | Menghapus checkpoint dan mengulang proses dari awal | `python bot_fasih_perbaikan.py --reset-checkpoint` |

---

## 📂 Struktur Direktori Proyek

```text
automated-fasih-sm/
│
├── .gitignore                   # Konfigurasi file yang diabaikan Git (.venv, output/, dll)
├── requirements.txt             # Daftar pustaka Python yang dibutuhkan
├── README.md                    # Dokumentasi lengkap proyek
├── bot_fasih_perbaikan.py       # Script inti bot otomasi Selenium
│
├── run_bot.bat                  # Shortcut menjalankan bot normal
├── retry_errors.bat             # Shortcut menjalankan bot khusus data error
├── test_url.bat                 # Shortcut menjalankan bot untuk 1 URL uji coba
├── start_chrome_ozifazh.bat     # Shortcut membuka Chrome dengan remote debugging port 9222
│
├── data/                        # Folder penyimpanan file sumber data Excel
│   └── 030-fix-rozi.xlsx
│
└── output/                      # Folder output hasil proses (otomatis dibuat)
    ├── 030-fix-rozi_hasil.xlsx  # Excel hasil dengan status eksekusi bot
    ├── checkpoint_*.json        # Database progres pengerjaan
    └── errors/                  # Folder tangkapan layar jika terjadi kendala
```

---

## ⚠️ Troubleshooting & Catatan Penting

1. **Koneksi Chrome Gagal (`GAGAL MENGHUBUNGKAN KE CHROME`)**:
   - Pastikan Google Chrome sudah dibuka melalui `start_chrome_ozifazh.bat`.
   - Pastikan tidak ada proses Chrome biasa yang memblokir port `9222`.
2. **Halaman Tidak Dapat Diakses (*This site can't be reached*)**:
   - Periksa koneksi **VPN BPS** Anda. Pastikan VPN dalam keadaan terhubung.
3. **Terkena Bot Detection**:
   - Biarkan bot berjalan; bot akan otomatis menghitung mundur 5 menit pendinginan lalu mengeklik tombol *Kembali*.
   - Jangan menjalankan bot secara agresif atau menurunkan jeda antar-request terlalu rendah.
