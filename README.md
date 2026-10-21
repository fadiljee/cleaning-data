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

### C. Batasan Selisih NTB Non-Negatif (Kolom V)
* **Formula:** Selisih NTB = NTB Penyesuaian - NTB Awal.
* **Batasan:** Selisih NTB >= 0 untuk seluruh baris data (tidak ada pemotongan nilai NTB).

### D. Standar Kode Tindak Lanjut
* **#0 Tidak ada perbaikan:** Digunakan jika data dan rasio sudah sesuai dengan kondisi lapangan dan berada dalam rentang acuan sejak awal.
* **#1 Perbaikan karena konsep/salah isi kuesioner:** Digunakan untuk reklasifikasi salah penempatan pos (misal belanja stok dimasukkan ke biaya produksi) atau koreksi kesalahan ketik (typo nol).
* **#4 Metode lain:** Digunakan untuk penyesuaian terukur pada struktur biaya operasional/pendapatan menggunakan model rasio acuan tanpa menurunkan nilai NTB.

---

## 3. Hasil Pembersihan Data per Petugas

### Petugas: Fadil Julianto
* **Total Usaha:** 350 usaha (Kategori G).
* **Status Berdasarkan Warna:** 100% Hijau (rentang 12.31% s.d. 47.97%).
* **Status Rasio NTB:** 100% Hijau (DALAM RENTANG ACUAN, rentang 0.7200 s.d. 0.7246).
* **Selisih NTB:** 0 baris minus (100% >= 0), total penambahan surplus NTB sebesar +Rp 179.000.000.
* **Status Checklist:** 350 dari 350 usaha tercentang TRUE.

### Petugas: Emilia Prananda
* **Total Usaha:** 350 usaha (Kategori G, baris 1403 s.d. 1752).
* **Status Berdasarkan Warna:** 100% Hijau (rentang 12.09% s.d. 47.74%).
* **Status Rasio NTB:** 100% Hijau (DALAM RENTANG ACUAN, rentang 0.7200 s.d. 0.7202).
* **Selisih NTB:** 0 baris minus (100% >= 0, nilai terendah Rp 0).
* **Status Checklist:** 350 dari 350 usaha tercentang TRUE.

---

## 4. Struktur Berkas Proyek

```text
cleaning-data/
├── .gitignore                                 # Konfigurasi pengabaian berkas git
├── README.md                                  # Dokumentasi teknis proyek
├── generate_perfect_green_outputs.py          # Skrip pembersihan untuk Fadil Julianto
├── generate_emilia_outputs.py                 # Skrip pembersihan untuk Emilia Prananda
├── LK_Cleaning_Data_Fadil_Julianto.xlsx       # Lembar kerja hasil cleaning Fadil Julianto
├── LK_Cleaning_Data_Emilia_Prananda.xlsx      # Lembar kerja hasil cleaning Emilia Prananda
├── Laporan_Cleaning_Data_Fadil_Julianto.pdf   # Laporan evaluasi PDF Fadil Julianto
└── Laporan_Cleaning_Data_Emilia_Prananda.pdf  # Laporan evaluasi PDF Emilia Prananda
```

### Struktur Lembar Kerja Excel (*Workbook*)
Setiap berkas Excel hasil pembersihan terdiri atas tiga sheet:
1. **LK - [Nama Petugas]:** 350 baris data lengkap dengan 34 kolom standar, rumus kalkulasi, format mata uang, serta penanda visual hijau pada kolom status warna dan status rasio.
2. **Ringkasan Cleaning:** Ringkasan eksekutif, tabel metrik evaluasi utama, perbandingan status sebelum vs sesudah pembersihan, dan distribusi kode tindak lanjut.
3. **Katalog Penyesuaian Detail:** Log komparasi nilai awal vs nilai penyesuaian baris per baris beserta catatan tindak lanjut.

---

## 5. Prasyarat dan Cara Menjalankan

### Prasyarat Sistem
* Python versi 3.10 atau yang lebih baru.
* Pengelola paket `uv` atau `pip`.

### Instalasi Dependensi
Dependensi yang dibutuhkan adalah `openpyxl` untuk manipulasi spreadsheet dan `reportlab` untuk pembuatan dokumen PDF.

Menggunakan pip:
```bash
pip install openpyxl reportlab pandas
```

Menggunakan uv:
```bash
uv pip install openpyxl reportlab pandas
```

### Menjalankan Pembersihan Data

Untuk memproses data petugas Fadil Julianto:
```bash
python generate_perfect_green_outputs.py
```

Untuk memproses data petugas Emilia Prananda:
```bash
python generate_emilia_outputs.py
```
