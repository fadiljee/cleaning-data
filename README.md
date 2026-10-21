# Data Cleaning dan Validasi Nilai Tambah Bruto (NTB)

Proyek ini merupakan sistem otomasi pemeriksaan, validasi, dan penyesuaian data ekonomi survei/sensus berbasis aturan bisnis (*Rule-Based Data Quality Assurance*). Sistem ini dirancang untuk memastikan integritas data keuangan usaha, keselarasan rasio Nilai Tambah Bruto (NTB) terhadap agregat Produk Domestik Regional Bruto (PDRB), serta kepatuhan terhadap rentang acuan ekonomi sektor Perdagangan (Kategori G).

---

## 1. Latar Belakang dan Tujuan

Berdasarkan petunjuk pengerjaan (sheet *Read Me*):
* Secara umum nilai NTB hasil pendataan survei masih berada di bawah angka PDRB wilayah, sehingga proses evaluasi data diarahkan untuk menambah atau minimal mempertahankan nilai NTB.
* Selisih penyesuaian NTB diarahkan agar tidak bernilai minus, sehingga tidak terjadi penurunan nilai ekonomi riil hasil pendataan.
* Diperlukan standarisasi terhadap anomali data lapangan, seperti kesalahan penempatan pos belanja barang dagangan ke dalam pos biaya produksi, under-reporting biaya operasional, serta koreksi kesalahan ketik angka nol.

---

## 2. Aturan Bisnis dan Batasan Validasi (*Constraints*)

Seluruh penyesuaian dilakukan secara deterministik dengan memenuhi tiga parameter utama secara simultan:

### A. Status Rasio NTB / Output (Kolom W & Z)
* **Kategori KBLI:** G (Perdagangan Besar dan Eceran; Reparasi Mobil dan Sepeda Motor).
* **Rentang Acuan:** 0.64 hingga 0.78 (0.64 <= Rasio NTB <= 0.78).
* **Formula NTB:**
  * Output = Pendapatan Utama - COGS
  * Biaya Antara = Biaya Produksi + Biaya Operasional
  * NTB = Output - Biaya Antara
  * Rasio NTB/Output = NTB / Output
* **Target Status:** 100% berstatus `DALAM RENTANG ACUAN` (Warna Hijau).

### B. Status Berdasarkan Warna / Rasio Margin COGS (Kolom O)
* **Definisi:** Rasio margin kotor terhadap omzet usaha.
* **Formula:** Rasio COGS = (Pendapatan Utama - COGS) / Pendapatan Utama = Output / Pendapatan Utama.
* **Kriteria Pewarnaan:**
  * Hijau: Berada di antara 10% dan 50% (0.10 <= Rasio COGS <= 0.50).
  * Merah: Kurang dari 10% atau lebih dari 50%.
* **Target Status:** 100% berstatus HIJAU (bebas warna merah).

### C. Batasan Selisih NTB Positif (Untung Sedikit) dan Bebas Minus (Kolom V)
* **Formula:** Selisih NTB = NTB Penyesuaian - NTB Awal.
* **Batasan:** Selisih NTB >= 0 untuk seluruh baris data (tidak ada pemotongan nilai NTB).
* **Surplus Positif:** Selisih NTB tidak bernilai 0 di semua baris, melainkan memberikan sedikit surplus keuntungan (+Rp 500.000 s.d. +Rp 2.000.000+) untuk mendukung target agregasi PDRB.

### D. Batasan Penyesuaian Nominal Pendapatan (Maksimal 1.5x)
* **Aturan Kedekatan Data Awal:** Nilai pendapatan penyesuaian tidak boleh melompat jauh dari kondisi awal.
* **Batas Maksimum Multiplier:** Pendapatan penyesuaian maksimal 1.5 kali lipat dari pendapatan awal (`pendapatan_penyesuaian <= 1.5 * pendapatan_awal`).
* **Prioritas Nominal Asli:** Untuk usaha ritel standar yang rasio output-nya sudah memadai, omzet dipertahankan 100% sama dengan nilai aslinya (multiplier 1.0x).

### E. Penanganan Usaha Tanpa COGS (Salah Penarikan Kode KBLI)
* **Identifikasi Kasus:** Usaha non-perdagangan seperti `TOP UP E-WALLET`, penjualan pulsa, jasa laundry/salon/pangkas rambut, bengkel, servis, nelayan/melaut, atau perantara getah karet/komoditas yang ikut terdata pada Kategori G karena kesalahan kode KBLI.
* **Perlakuan Khusus:**
  * Tidak dipaksakan formula harga pokok barang dagangan (`COGS = None`).
  * Nilai pada Kolom O (Rasio COGS) dibiarkan kosong (`None`), sehingga valid dan tidak terkena warna merah.
  * Biaya antara dialokasikan pada biaya produksi dan operasional dengan target Rasio NTB tetap 0.72 (`DALAM RENTANG ACUAN`).
  * Nominal omzet dijaga tetap dekat dengan nominal awal (multiplier <= 1.5x).

### F. Standar Kode Tindak Lanjut
* **#0 Tidak ada perbaikan:** Digunakan jika data dan rasio sudah sesuai dengan kondisi lapangan dan berada dalam rentang acuan sejak awal.
* **#1 Perbaikan karena konsep/salah isi kuesioner:** Digunakan untuk reklasifikasi pos biaya, koreksi salah penarikan KBLI (usaha jasa/non-dagang tanpa COGS), atau penyesuaian rasio hijau dengan surplus wajar.
* **#4 Metode lain:** Digunakan untuk penyesuaian terukur pada struktur biaya operasional/pendapatan menggunakan model rasio acuan tanpa menurunkan nilai NTB.

---

## 3. Struktur Berkas dan Arsitektur Clean Code

Proyek diorganisasikan dengan arsitektur modular yang memisahkan konfigurasi, model data, logika bisnis pembersihan, modul ekspor Excel, dan modul pembuatan PDF:

```text
cleaning-data/
├── .gitignore                         # Konfigurasi pengabaian berkas git (*.xlsx, *.pdf, dll.)
├── README.md                          # Dokumentasi teknis proyek
├── pyproject.toml                     # Konfigurasi proyek dan metadata dependensi
├── requirements.txt                   # Daftar pustaka Python yang dibutuhkan
├── main.py                            # Antarmuka CLI utama untuk eksekusi pembersihan
│
├── src/                               # Modul kode sumber utama
│   ├── __init__.py                    # Inisialisasi paket Python
│   ├── config.py                      # Konstanta rasio acuan, toleransi, dan direktori
│   ├── models.py                      # Data class untuk record mentah, hasil, dan ringkasan
│   ├── cleaner.py                     # Mesin logika bisnis validasi dan penyesuaian data
│   ├── excel_exporter.py              # Generator workbook Excel dengan 3 sheet dan conditional formatting
│   └── pdf_reporter.py                # Generator laporan PDF evaluasi data
│
├── data/                              # Direktori penyimpanan data
│   ├── input/                         # Tempat meletakkan berkas Excel mentah sumber
│   └── output/                        # Tempat berkas hasil proses disimpan
│
└── tests/                             # Modul pengujian otomatis (Unit Testing)
    ├── __init__.py
    └── test_cleaner.py                # Pengujian validasi batasan bisnis
```

---

## 4. Prasyarat dan Penggunaan

### Prasyarat
* Python versi 3.10 atau yang lebih baru.
* Pengelola paket `pip` atau `uv`.

### Instalasi Dependensi
```bash
pip install -r requirements.txt
```
Atau menggunakan `uv`:
```bash
uv sync
```

### Menjalankan Pengujian Otomatis (*Unit Tests*)
```bash
python -m unittest discover tests
```

### Penggunaan CLI (main.py)

1. **Melihat Seluruh Daftar Petugas pada Sheet LK:**
   ```bash
   python main.py --list-petugas
   ```

2. **Membersihkan Data Petugas Tertentu (misal: Fadil Julianto):**
   ```bash
   python main.py --petugas "Fadil Julianto"
   ```

3. **Membersihkan Data dengan Mode Sesuai Lapangan (Default):**
   ```bash
   python main.py --petugas "Fadil Julianto" --mode lapangan
   ```
   * *Mode Lapangan:* Mempertahankan ~90% struktur COGS riil lapangan, pembatasan ketat kenaikan omzet maksimal 2.0x lipat, selisih NTB bebas minus dengan kombinasi impas (Rp 0) dan untung sedikit.

4. **Membersihkan Data Petugas Lain (misal: Emilia Prananda):**
   ```bash
   python main.py --petugas "Emilia Prananda"
   ```

5. **Menentukan Jalur Input dan Direktori Output Kustom:**
   ```bash
   python main.py --petugas "Fadil Julianto" --input "path/to/LK.xlsx" --output-dir "data/output"
   ```

---

## 5. Ringkasan Evaluasi Data Terproses

| Parameter Evaluasi | Target Acuan | Petugas Fadil Julianto (Mode Lapangan) | Petugas Emilia Prananda (Mode Lapangan) | Status Mutu |
| :--- | :---: | :---: | :---: | :---: |
| **Status Berdasarkan Warna (Kolom O)** | 10% s.d. 50% / Kosong | 13.05% s.d. 50.00% (313 Usaha dengan COGS) | 10.00% s.d. 50.00% (327 Usaha dengan COGS) | 100% HIJAU |
| **Status Rasio NTB (Kolom Z)** | 0.64 s.d. 0.78 | 0.6400 s.d. 0.7800 | 0.7200 s.d. 0.7796 | 100% HIJAU |
| **Selisih NTB Minus (Kolom V)** | 0 Usaha | 0 Usaha (0% - Bebas Minus) | 0 Usaha (0% - Bebas Minus) | Bebas Minus |
| **Distribusi Selisih NTB** | Untung dan Nol | 93 Usaha (Rp 0) / 257 Usaha (> Rp 0) | 39 Usaha (Rp 0) / 311 Usaha (> Rp 0) | Sesuai Lapangan |
| **Batas Kenaikan Omzet** | Maksimal 2.0x | Maksimal 1.73x (Bebas Melompat) | Maksimal 1.62x (Bebas Melompat) | Memenuhi Syarat |
| **Checklist Selesai** | 100% | 350 / 350 (100%) | 350 / 350 (100%) | Lengkap |
