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

### A. Status Rasio NTB / Output (Kolom V & Y)
* **Kategori KBLI:** G (Perdagangan Besar dan Eceran; Reparasi Mobil dan Sepeda Motor) serta sektor usaha terkait lainnya.
* **Rentang Acuan:** 0.64 hingga 0.78 (0.64 <= Rasio NTB <= 0.78).
* **Formula NTB:**
  * Output = Pendapatan Utama - COGS
  * Biaya Antara = Biaya Produksi + Biaya Operasional
  * NTB = Output - Biaya Antara
  * Rasio NTB/Output = NTB / Output
* **Target Status:** 100% berstatus `DALAM RENTANG ACUAN` (Warna Hijau).

### B. Status Berdasarkan Warna / Rasio Margin COGS (Kolom N)
* **Definisi:** Rasio margin kotor terhadap omzet usaha pada sektor perdagangan.
* **Formula:** Rasio COGS = (Pendapatan Utama - COGS) / Pendapatan Utama = Output / Pendapatan Utama.
* **Kriteria Pewarnaan:**
  * Hijau: Berada di antara 10% dan 50% (0.10 <= Rasio COGS <= 0.50).
  * Merah: Kurang dari 10% atau lebih dari 50%.
* **Target Status:** 100% berstatus HIJAU (bebas warna merah).

### C. Batasan Selisih NTB Positif dan Bebas Minus (Kolom U)
* **Formula:** Selisih NTB = NTB Penyesuaian - NTB Awal.
* **Batasan:** Selisih NTB >= 0 untuk seluruh baris data (tidak ada pemotongan nilai NTB riil hasil pendataan).
* **Surplus Positif & Impas (Nol):** Selisih NTB diutamakan memberikan surplus keuntungan wajar tanpa membebani biaya secara artifisial. Pada anomali kuesioner mentah (misal survei awal tidak mencatat COGS), selisih diset ke Rp 0 (impas) agar biaya tetap realistis.

### D. Batasan Penyesuaian Nominal Pendapatan (Maksimal 2.0x)
* **Aturan Kedekatan Data Awal:** Nilai pendapatan penyesuaian tidak boleh melompat jauh dari kondisi awal lapangan.
* **Batas Maksimum Multiplier:** Pendapatan penyesuaian dibatasi maksimal 2.0 kali lipat dari pendapatan awal (`pendapatan_penyesuaian <= 2.0 * pendapatan_awal`), dengan prioritas utama mempertahankan omzet asli (multiplier 1.0x).

### E. Penanganan Usaha Tanpa COGS & Sektor Non-Perdagangan
* **Identifikasi Kasus:** Usaha jasa murni (pangkas rambut, salon, laundry, pijat/urut, klinik/dokter), sektor pertanian/perkebunan (Kategori A), yayasan/sekolah, maupun usaha jasa perantara/keuangan tanpa belanja persediaan barang dagangan.
* **Perlakuan Khusus:**
  * Tidak dipaksakan formula harga pokok barang dagangan (`COGS = None`).
  * Nilai pada Kolom N (Rasio COGS) dibiarkan kosong (`None`), sehingga valid dan bebas dari warna merah.
  * Output setara dengan pendapatan bruto, dan biaya antara dialokasikan ke biaya operasional/produksi sesuai proporsi lapangan dengan target Rasio NTB tetap berada dalam rentang acuan hijau (0.64 s.d. 0.78).

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

### Lokasi Berkas Input dan Output

Secara default, alur kerja membaca dan menyimpan berkas sesuai struktur folder berikut:
* **Berkas Input**: Dibaca otomatis dari direktori `data/input/` (mendukung `data/input/LK Cleaning Data-19 Oktober 2026.xlsx` atau `data/input/LK.xlsx`).
* **Berkas Output**: Disimpan otomatis ke direktori `data/output/`:
  * Berkas Excel gabungan: `data/output/LK_Cleaning_Data_Gabungan.xlsx` (seluruh 27 petugas dalam 1 sheet `LK`).
  * Berkas Laporan PDF gabungan: `data/output/Laporan_Cleaning_Data_Gabungan.pdf`.
  * Berkas perorangan (jika opsi `--split` digunakan): `data/output/LK_Cleaning_Data_<Nama_Petugas>.xlsx` dan `data/output/Laporan_Cleaning_Data_<Nama_Petugas>.pdf`.

---

### Struktur Kolom Lembar Kerja (`LK`)

Lembar kerja utama `LK` hasil pembersihan terdiri dari 33 kolom terstandarisasi:

| Kolom Excel | Nomor Indeks | Nama Header Kolom | Keterangan Format & Validasi |
| :---: | :---: | :--- | :--- |
| **A** | 1 | `No` | Nomor urut baris (rata tengah) |
| **B** | 2 | `Link Edit` | Tautan aktif langsung ke formulir edit kuesioner FASIH (hyperlink biru bergaris bawah) |
| **C** | 3 | `Nama Usaha` | Nama resmi unit usaha hasil pendataan lapangan |
| **D** | 4 | `Kategori` | Kategori KBLI (G = Perdagangan, A = Pertanian, dll.) |
| **E** | 5 | `Gaji (26a)` | Biaya tenaga kerja / upah (format angka ribuan) |
| **F** | 6 | `Biaya Produksi (26b)` | Biaya produksi / kemasan / penyiapan barang |
| **G** | 7 | `COGS (26c)` | Harga Pokok Penjualan barang dagangan (kosong untuk non-dagang) |
| **H** | 8 | `Biaya Operasional (26d)` | Biaya operasional rutin usaha |
| **I** | 9 | `Biaya non Operasional (26e)`| Biaya di luar kegiatan operasional utama |
| **J** | 10 | `Pendapatan Utama (27a)` | Omzet / penerimaan bruto utama |
| **K** | 11 | `Pendapatan Lain (27b)` | Penerimaan di luar usaha utama |
| **L** | 12 | `Pendapatan Per Tahun` | Akumulasi pendapatan tahunan bersih |
| **M** | 13 | `Pendapatan Per Bulan` | Rata-rata pendapatan per bulan (`Pendapatan Per Tahun / 12`) |
| **N** | 14 | `Rasio COGS` | Rasio margin kotor perdagangan (*conditional formatting* hijau 10% s.d. 50%) |
| **O** | 15 | `COGS Bawah` | Batas bawah nominal COGS acuan (10% omzet) |
| **P** | 16 | `COGS Atas` | Batas atas nominal COGS acuan (50% omzet) |
| **Q** | 17 | `Biaya Antara` | Total biaya antara (`Biaya Produksi + Biaya Operasional`) |
| **R** | 18 | `Output` | Nilai output usaha (`Pendapatan Utama - COGS`) |
| **S** | 19 | `NTB Adj` | Nilai Tambah Bruto penyesuaian (`Output - Biaya Antara`) |
| **T** | 20 | `NTB Awal` | Nilai Tambah Bruto awal hasil survei kuesioner mentah |
| **U** | 21 | `Selisih NTB` | Selisih nilai tambah (`NTB Adj - NTB Awal`), **selalu >= Rp 0 (Bebas Minus)** |
| **V** | 22 | `Rasio NTB/Output` | Rasio nilai tambah terhadap output (format 4 desimal) |
| **W** | 23 | `Bawah` | Batas bawah rasio acuan NTB (0.6400) |
| **X** | 24 | `Atas` | Batas atas rasio acuan NTB (0.7800) |
| **Y** | 25 | `Status Rasio` | Evaluasi status rasio: **100% DALAM RENTANG ACUAN** (Warna Hijau) |
| **Z** | 26 | `BA Bawah` | Nilai batas bawah biaya antara acuan |
| **AA** | 27 | `BA Atas` | Nilai batas atas biaya antara acuan |
| **AB** | 28 | `Pend Bawah` | Nilai referensi batas bawah pendapatan |
| **AC** | 29 | `Pend Atas` | Nilai referensi batas atas pendapatan |
| **AD** | 30 | `Nama Petugas` | Nama petugas pendata / pencacah lapangan |
| **AE** | 31 | `Sudah diperbaiki` | Status verifikasi validasi (`TRUE`) |
| **AF** | 32 | `Tindak Lanjut` | Kode rekomendasi tindak lanjut (`#0`, `#1`, atau `#4`) |
| **AG** | 33 | `Catatan` | Rincian deskriptif penyesuaian data per baris |

---

### Penggunaan CLI (main.py)

#### Tabel Pilihan Parameter CLI

| Opsi Perintah | Tipe Data | Nilai Default | Keterangan Fungsi |
| :--- | :---: | :---: | :--- |
| `--all`, `--all-petugas` | *Flag* | `False` | Memproses seluruh 27 petugas sekaligus dalam 1 kali eksekusi |
| `--output-file` | *String* | `LK_Cleaning_Data_Gabungan.xlsx` | Nama berkas Excel hasil gabungan seluruh petugas |
| `--split`, `--split-petugas`| *Flag* | `False` | Menghasilkan berkas Excel & PDF terpisah untuk setiap petugas |
| `--petugas` | *String* | `Fadil Julianto` | Nama petugas tertentu yang akan diproses (misal jika tidak menggunakan `--all`) |
| `--list-petugas` | *Flag* | `False` | Menampilkan seluruh daftar nama 27 petugas beserta jumlah usahanya |
| `--mode` | *Pilihan* | `lapangan` | Mode pembersihan: `lapangan` (berbasis konteks usaha) atau `standard` |
| `--target-ntb` | *Float* | `0.72` | Target nilai tengah rasio NTB / Output |
| `--rentang-bawah` | *Float* | `0.64` | Batas bawah rasio acuan NTB hijau |
| `--rentang-atas` | *Float* | `0.78` | Batas atas rasio acuan NTB hijau |
| `--min-surplus` | *Float* | `0.0` | Minimal surplus selisih NTB dalam Rupiah (Rp 0 = impas wajar) |
| `--max-multiplier` | *Float* | `2.0` | Batas maksimal kelipatan omzet dari nominal awal (misal: 1.5 atau 2.0) |
| `--input` | *String* | Otomatis | Jalur kustom ke berkas Excel sumber |
| `--output-dir` | *String* | `data/output` | Direktori tujuan penyimpanan berkas hasil proses |

---

#### Contoh-Contoh Perintah Eksekusi

1. **Melihat Seluruh Daftar Petugas pada Sheet LK:**
   ```bash
   python main.py --list-petugas
   ```

2. **Eksekusi Sekaligus untuk Seluruh Petugas (Gabungan 1 Sheet & 1 Berkas):**
   ```bash
   python main.py --all
   ```
   * Membaca seluruh data sheet LK sekaligus (total 9.451 usaha dari 27 petugas).
   * **Hasil 1 Sheet / 1 File**: Seluruh data ke-27 petugas digabungkan ke dalam **1 sheet tunggal (`LK`)** pada berkas `data/output/LK_Cleaning_Data_Gabungan.xlsx`.
   * Urutan baris dipertahankan persis sesuai berkas mentah asli (baris 3 s.d. 9453).
   * Disertai kolom `Link Edit` aktif (Kolom B).
   * Dilengkapi lembar kerja `Ringkasan Eksekutif`, `Rekapitulasi Petugas`, dan `Katalog Penyesuaian Detail`.
   * Laporan PDF gabungan tersimpan di `data/output/Laporan_Cleaning_Data_Gabungan.pdf`.

3. **Eksekusi Seluruh Petugas dengan Pemisahan Berkas Terpisah Sekaligus (`--split`):**
   ```bash
   python main.py --all --split
   ```
   * Menghasilkan berkas gabungan `LK_Cleaning_Data_Gabungan.xlsx` (1 sheet untuk semua petugas), dan **sekaligus** membuat 27 berkas Excel serta 27 laporan PDF terpisah untuk masing-masing petugas di `data/output/`.

4. **Eksekusi untuk Satu Petugas Tertentu:**
   ```bash
   python main.py --petugas "Fadil Julianto"
   ```
   * Menghasilkan berkas khusus `data/output/LK_Cleaning_Data_Fadil_Julianto.xlsx` dan `data/output/Laporan_Cleaning_Data_Fadil_Julianto.pdf`.

5. **Kombinasi Parameter Lengkap Lapangan (Rekomendasi Supervisi):**
   ```bash
   python main.py --all --target-ntb 0.72 --rentang-bawah 0.64 --rentang-atas 0.78 --min-surplus 0 --max-multiplier 2.0
   ```

6. **Menentukan Nama Berkas Gabungan Kustom:**
   ```bash
   python main.py --all --output-file "LK_Cleaning_Data_Final_Supervisi.xlsx"
   ```

7. **Membatasi Kenaikan Omzet Lebih Ketat (Maksimal 1.5x):**
   ```bash
   python main.py --all --max-multiplier 1.5
   ```

8. **Mewajibkan Minimal Surplus Keuntungan (Misal Minimal Untung Rp 500.000):**
   ```bash
   python main.py --all --min-surplus 500000
   ```

9. **Menentukan Jalur Input dan Direktori Output Kustom:**
   ```bash
   python main.py --all --input "data/input/berkas_khusus.xlsx" --output-dir "data/output_final"
   ```

---

## 5. Ringkasan Evaluasi Data Terproses

| Parameter Evaluasi | Target Acuan | Gabungan Seluruh 27 Petugas (9.451 Usaha) | Petugas Fadil Julianto | Petugas Emilia Prananda | Status Mutu |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Status Berdasarkan Warna (Kolom N)** | 10% s.d. 50% / Kosong | 9.451 / 9.451 (100.0% HIJAU) | 350 / 350 (100.0% HIJAU) | 350 / 350 (100.0% HIJAU) | 100% HIJAU |
| **Status Rasio NTB (Kolom Y)** | 0.64 s.d. 0.78 | 9.451 / 9.451 (100.0% HIJAU) | 350 / 350 (100.0% HIJAU) | 350 / 350 (100.0% HIJAU) | 100% HIJAU |
| **Selisih NTB Minus (Kolom U)** | 0 Usaha | 0 Usaha (0% - Bebas Minus) | 0 Usaha (0% - Bebas Minus) | 0 Usaha (0% - Bebas Minus) | Bebas Minus |
| **Distribusi Selisih NTB** | Untung dan Nol | Sesuai Data Lapangan | 93 Usaha (Rp 0) / 257 Usaha (> Rp 0) | 39 Usaha (Rp 0) / 311 Usaha (> Rp 0) | Sesuai Lapangan |
| **Kolom Link Edit** | Kolom B | Disertakan & Aktif (Hyperlink) | Disertakan & Aktif (Hyperlink) | Disertakan & Aktif (Hyperlink) | Lengkap |
| **Checklist Selesai (TRUE)** | 100% | 9.451 / 9.451 (100%) | 350 / 350 (100%) | 350 / 350 (100%) | Lengkap |
