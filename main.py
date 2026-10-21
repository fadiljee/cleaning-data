"""
Entrypoint CLI Utama untuk Pembersihan dan Validasi Data LK
"""
import argparse
import sys
from pathlib import Path

import openpyxl

from src.cleaner import DataCleaner
from src.cleaner_lapangan import DataCleanerLapangan
from src.excel_exporter import ExcelExporter
from src.pdf_reporter import PDFReporter
from src.config import BASE_DIR, DEFAULT_INPUT_FILE


def find_input_file(custom_path: str = None) -> Path:
    """Mencari berkas input Excel di beberapa lokasi standar"""
    if custom_path:
        p = Path(custom_path)
        if p.exists():
            return p
        raise FileNotFoundError(f"Berkas input tidak ditemukan: {custom_path}")

    # Cek di data/input lalu root project
    search_paths = [
        BASE_DIR / "data" / "input" / DEFAULT_INPUT_FILE,
        BASE_DIR / DEFAULT_INPUT_FILE,
    ]
    for p in search_paths:
        if p.exists():
            return p

    # Cari file xlsx apapun yang diawali LK Cleaning Data
    for candidate in BASE_DIR.glob("LK Cleaning Data*.xlsx"):
        return candidate
    for candidate in (BASE_DIR / "data" / "input").glob("LK Cleaning Data*.xlsx"):
        return candidate

    raise FileNotFoundError(f"Tidak dapat menemukan berkas '{DEFAULT_INPUT_FILE}'.")


def list_available_petugas(workbook_path: Path):
    """Menampilkan daftar seluruh nama petugas yang terdapat pada sheet LK"""
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    if "LK" not in wb.sheetnames:
        print("Sheet 'LK' tidak ditemukan.")
        return

    sheet = wb["LK"]
    petugas_counts = {}
    for row in sheet.iter_rows(min_row=3, values_only=True):
        p = row[30]
        if p:
            p_str = str(p).strip()
            petugas_counts[p_str] = petugas_counts.get(p_str, 0) + 1

    print("\n=== DAFTAR PETUGAS PADA SHEET LK ===")
    print(f"{'No':<4} | {'Nama Petugas':<45} | {'Jumlah Usaha':<12}")
    print("-" * 66)
    for idx, (p, cnt) in enumerate(sorted(petugas_counts.items(), key=lambda x: x[0]), start=1):
        print(f"{idx:<4} | {p:<45} | {cnt:<12}")
    print("-" * 66)
    print(f"Total Petugas: {len(petugas_counts)}")


def process_petugas(cleaner: DataCleaner, petugas_name: str, output_dir: Path):
    """Menjalankan alur lengkap pembersihan dan pembuatan berkas untuk satu petugas"""
    clean_name = petugas_name.replace(" ", "_").replace(".", "").replace(",", "")
    safe_name = "".join(c for c in clean_name if c.isalnum() or c == "_")

    print(f"\n[INFO] Membaca data untuk petugas: '{petugas_name}'...")
    raw_records = cleaner.load_records_by_petugas(petugas_name)

    if not raw_records:
        print(f"[PERINGATAN] Tidak ditemukan data untuk petugas '{petugas_name}'.")
        return

    print(f"[INFO] Ditemukan {len(raw_records)} usaha. Menjalankan algoritma pembersihan data...")
    cleaned_records, summary = cleaner.clean_records(raw_records)

    # Nama berkas output
    out_xlsx = output_dir / f"LK_Cleaning_Data_{safe_name}.xlsx"
    out_pdf = output_dir / f"Laporan_Cleaning_Data_{safe_name}.pdf"

    print(f"[INFO] Membuat berkas Excel: {out_xlsx.name}...")
    exporter = ExcelExporter(str(out_xlsx))
    exporter.export(cleaned_records, summary)

    print(f"[INFO] Membuat laporan PDF: {out_pdf.name}...")
    reporter = PDFReporter(str(out_pdf))
    reporter.generate(cleaned_records, summary)

    print("\n=== HASIL EVALUASI AKHIR ===")
    print(f"Petugas               : {summary.petugas}")
    print(f"Total Usaha           : {summary.total_usaha}")
    print(f"Status Warna Hijau    : {summary.total_warna_hijau}/{summary.total_usaha} (100% HIJAU, rentang {summary.min_cogs_ratio*100:.2f}% s.d. {summary.max_cogs_ratio*100:.2f}%)")
    print(f"Status Rasio NTB      : {summary.total_rasio_hijau}/{summary.total_usaha} (100% HIJAU DALAM ACUAN, rentang {summary.min_ntb_ratio:.4f} s.d. {summary.max_ntb_ratio:.4f})")
    print(f"Selisih NTB Minus     : {summary.total_selisih_minus} (0% - TIDAK ADA MINUS)")
    print(f"Total Surplus NTB     : Rp {summary.total_selisih_ntb:,.0f}")
    print(f"Berkas Excel          : {out_xlsx.resolve()}")
    print(f"Berkas PDF            : {out_pdf.resolve()}")


def main():
    parser = argparse.ArgumentParser(
        description="Sistem Otomasi Data Cleaning dan Validasi NTB Kategori G"
    )
    parser.add_argument(
        "--petugas",
        type=str,
        help="Nama petugas yang akan dibersihkan datanya (misal: 'Fadil Julianto' atau 'Emilia Prananda')"
    )
    parser.add_argument(
        "--list-petugas",
        action="store_true",
        help="Menampilkan seluruh daftar nama petugas yang ada di file Excel"
    )
    parser.add_argument(
        "--input",
        type=str,
        help="Jalur khusus ke berkas Excel input (opsional)"
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=".",
        help="Direktori penyimpanan hasil output Excel dan PDF (default: folder saat ini)"
    )
    parser.add_argument(
        "--mode",
        type=str,
        choices=["lapangan", "standard"],
        default="lapangan",
        help="Mode pembersihan: 'lapangan' (90% sesuai lapangan, max 2x) atau 'standard'"
    )

    args = parser.parse_args()

    try:
        input_file = find_input_file(args.input)
    except FileNotFoundError as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)

    if args.list_petugas:
        list_available_petugas(input_file)
        sys.exit(0)

    if not args.petugas:
        print("[INFO] Parameter --petugas tidak diberikan. Menjalankan default: 'Fadil Julianto'")
        petugas_target = "Fadil Julianto"
    else:
        petugas_target = args.petugas

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.mode == "lapangan":
        print("[INFO] Menggunakan Mode 'Sesuai Lapangan' (max 2x lipat, selisih NTB >= 0)")
        cleaner = DataCleanerLapangan(str(input_file))
    else:
        print("[INFO] Menggunakan Mode 'Standard'")
        cleaner = DataCleaner(str(input_file))

    process_petugas(cleaner, petugas_target, output_dir)


if __name__ == "__main__":
    main()
