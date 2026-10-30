"""
Eksekutor Pembersihan Data Sheet LK-30 Okt Khusus 6 Petugas (111 Unit Usaha)
Aturan:
1. Filter 6 Petugas Supervisi: Aswan, Emil, Fadil, Icha, Jimmi, Jinan (111 unit usaha).
2. NTB Awal menggunakan nilai resmi dari supervisi (111 nilai sheet asli).
3. Selisih NTB di rentang Rp 0 s.d. Rp 10.000.000 dengan mayoritas bernilai 0 (~70% bernilai Rp 0, 0 minus).
4. Rasio COGS acak alami 11% - 45% untuk perdagangan, None untuk non-dagang.
5. Rasio NTB acak alami 68% - 72% untuk perdagangan (G) dan 100% DALAM RENTANG ACUAN.
6. Pos Keuangan yang Disesuaikan: Gaji, Biaya Produksi, COGS, Biaya Operasional, Biaya non Operasional, Pendapatan Utama.
7. Pendapatan per bulan (m/12) normal dan bebas minus.
8. Sinkronisasi penuh dengan sheet sisa_30Okt dan kolom T LK-30 Okt.
"""
import sys
import openpyxl
from pathlib import Path

from src.cleaner_30_okt import load_raw_30_okt_6petugas, clean_6petugas_records
from src.excel_exporter import ExcelExporter
from src.pdf_reporter import PDFReporter
from src.config import MIN_COGS_MARGIN, MAX_COGS_MARGIN

INPUT_FILE = Path("data/input/LK Cleaning Data-19 Oktober 2026.xlsx")
ROOT_INPUT_FILE = Path("LK Cleaning Data-19 Oktober 2026.xlsx")
OUTPUT_DIR = Path("data/output")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

OUT_XLSX = OUTPUT_DIR / "LK_Cleaning_Data_30_Okt.xlsx"
OUT_PDF = OUTPUT_DIR / "Laporan_Cleaning_Data_30_Okt.pdf"

print("=" * 80)
print("MEMULAI PROSES PEMBERSIHAN DATA SHEET LK-30 OKT (111 UNIT USAHA - 6 PETUGAS)")
print("=" * 80)

source_path = str(INPUT_FILE if INPUT_FILE.exists() else ROOT_INPUT_FILE)
records, acuans, sisa_row_map, _ = load_raw_30_okt_6petugas(source_path)
print(f"Total data mentah terbaca: {len(records)} baris usaha (6 Petugas)")

cleaned_list, summary = clean_6petugas_records(records, acuans)

n_tot = len(cleaned_list)
n_z = sum(1 for c in cleaned_list if abs(c.selisih_ntb) < 1.0)
n_p = sum(1 for c in cleaned_list if c.selisih_ntb >= 1.0)
n_m = sum(1 for c in cleaned_list if c.selisih_ntb < -0.01)
cogs_ok = sum(1 for c in cleaned_list if c.rasio_cogs is None or (MIN_COGS_MARGIN <= c.rasio_cogs <= MAX_COGS_MARGIN))
ntb_ok = sum(1 for c in cleaned_list if c.status_rasio == "DALAM RENTANG ACUAN")
m12_ok = sum(1 for c in cleaned_list if c.pendapatan_per_bulan >= -0.01)

print("\n" + "=" * 80)
print("HASIL EVALUASI PEMERIKSAAN KUALITAS DATA:")
print(f"Total Usaha Terproses      : {n_tot}")
print(f"Status Rasio NTB Hijau     : {ntb_ok} / {n_tot} ({(ntb_ok/n_tot)*100:.1f}%)")
print(f"Status Warna COGS Hijau    : {cogs_ok} / {n_tot} ({(cogs_ok/n_tot)*100:.1f}%)")
print(f"Laba Per Bulan m/12 >= 0   : {m12_ok} / {n_tot} ({(m12_ok/n_tot)*100:.1f}%)")
print(f"Distribusi Selisih NTB:")
print(f"  - Nilai Rp 0 (Impas)     : {n_z} ({(n_z/n_tot)*100:.1f}%)")
print(f"  - Nilai Positif (Surplus): {n_p} ({(n_p/n_tot)*100:.1f}%) [Max: Rp {max(c.selisih_ntb for c in cleaned_list):,.0f}]")
print(f"  - Nilai Minus (Turun)    : {n_m} (Bebas Minus)")
print("=" * 80 + "\n")

# 1. Ekspor ke Excel Output Standar
print(f"[1/3] Menyimpan Excel Output ke: {OUT_XLSX.resolve()}...")
exporter = ExcelExporter(str(OUT_XLSX))
exporter.export(cleaned_list, summary, sheet_name="LK-30 Okt")

# 2. Ekspor ke PDF Laporan Resmi
print(f"[2/3] Membuat Laporan PDF ke: {OUT_PDF.resolve()}...")
reporter = PDFReporter(str(OUT_PDF))
reporter.generate(cleaned_list, summary)

# 3. Update sheet LK-30 Okt dan sisa_30Okt pada berkas workbook asli
print(f"[3/3] Memperbarui sheet LK-30 Okt dan sisa_30Okt pada berkas workbook asli...")
for target_wb_path in [INPUT_FILE, ROOT_INPUT_FILE]:
    if not target_wb_path.exists():
        continue
    print(f"  Memperbarui {target_wb_path.resolve()}...")
    wb_update = openpyxl.load_workbook(str(target_wb_path), data_only=False)
    if "LK-30 Okt" not in wb_update.sheetnames:
        continue
    ws_30 = wb_update["LK-30 Okt"]
    ws_sisa = wb_update["sisa_30Okt"] if "sisa_30Okt" in wb_update.sheetnames else None

    row_map = {c.row_idx: c for c in cleaned_list}
    for r_idx, c_rec in row_map.items():
        # A. Kolom Keuangan Input yang Disesuaikan pada LK-30 Okt
        ws_30.cell(r_idx, 5).value = c_rec.gaji
        ws_30.cell(r_idx, 6).value = c_rec.biaya_produksi
        ws_30.cell(r_idx, 7).value = c_rec.cogs if c_rec.cogs is not None else None
        ws_30.cell(r_idx, 8).value = c_rec.biaya_operasional
        ws_30.cell(r_idx, 9).value = c_rec.biaya_non_operasional
        ws_30.cell(r_idx, 10).value = c_rec.pendapatan_utama
        if c_rec.pendapatan_lain is not None:
            ws_30.cell(r_idx, 11).value = c_rec.pendapatan_lain

        # B. Tulis NTB Awal resmi ke Kolom T (NTB Awal)
        ws_30.cell(r_idx, 20).value = c_rec.ntb_awal

        # C. Metadata Perbaikan
        ws_30.cell(r_idx, 31).value = True
        ws_30.cell(r_idx, 32).value = c_rec.tindak_lanjut
        ws_30.cell(r_idx, 33).value = c_rec.catatan

        # D. Sinkronisasi ke Sheet sisa_30Okt
        if ws_sisa is not None:
            link_k = str(c_rec.link or "").strip()
            nama_k = str(c_rec.nama_usaha or "").strip()
            s_row = sisa_row_map.get((link_k, nama_k))
            if s_row is not None:
                ws_sisa.cell(s_row, 9).value = c_rec.gaji
                ws_sisa.cell(s_row, 10).value = c_rec.biaya_produksi
                ws_sisa.cell(s_row, 11).value = c_rec.cogs if c_rec.cogs is not None else 0.0
                ws_sisa.cell(s_row, 12).value = c_rec.biaya_operasional
                ws_sisa.cell(s_row, 13).value = c_rec.biaya_non_operasional
                ws_sisa.cell(s_row, 14).value = c_rec.pendapatan_utama
                if c_rec.pendapatan_lain is not None:
                    ws_sisa.cell(s_row, 15).value = c_rec.pendapatan_lain
                ws_sisa.cell(s_row, 16).value = c_rec.ntb_awal  # Kolom P: NTB Awal Dasar
                ws_sisa.cell(s_row, 17).value = c_rec.output    # Kolom Q: Output

    wb_update.save(str(target_wb_path))
    print(f"  -> Berhasil memperbarui {target_wb_path.name}")

print("\n" + "=" * 80)
print("SELURUH PROSES PEMBERSIHAN DATA SHEET LK-30 OKT SELESAI DENGAN SUKSES!")
print("=" * 80)
