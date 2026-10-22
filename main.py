"""
Entrypoint CLI Utama untuk Pembersihan dan Validasi Data LK
"""
import argparse
import sys
import warnings
from pathlib import Path

# Abaikan peringatan format sel tanggal bawaan openpyxl dari file Excel BPS
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import openpyxl

from src.cleaner import DataCleaner
from src.cleaner_lapangan import DataCleanerLapangan
from src.excel_exporter import ExcelExporter
from src.pdf_reporter import PDFReporter
from src.models import CleaningSummary
from src.config import BASE_DIR, DEFAULT_INPUT_FILE, INPUT_DIR, OUTPUT_DIR, MIN_COGS_MARGIN, MAX_COGS_MARGIN


def find_input_file(custom_path: str = None) -> Path:
    """Mencari berkas input Excel di beberapa lokasi standar dengan prioritas data/input/"""
    if custom_path:
        p = Path(custom_path)
        if p.exists():
            return p
        raise FileNotFoundError(f"Berkas input tidak ditemukan: {custom_path}")

    # 1. Cek berkas spesifik di data/input/
    search_paths = [
        INPUT_DIR / DEFAULT_INPUT_FILE,
        INPUT_DIR / "LK.xlsx",
    ]
    for p in search_paths:
        if p.exists():
            return p

    # 2. Cari file .xlsx apapun yang valid di data/input/
    if INPUT_DIR.exists():
        for candidate in sorted(INPUT_DIR.glob("*.xlsx")):
            if not candidate.name.startswith("~") and not candidate.name.startswith("."):
                return candidate

    # 3. Fallback jika belum dipindah dari root project
    search_root = [
        BASE_DIR / DEFAULT_INPUT_FILE,
    ]
    for p in search_root:
        if p.exists():
            return p

    for candidate in BASE_DIR.glob("LK Cleaning Data*.xlsx"):
        if not candidate.name.startswith("~") and not candidate.name.startswith("."):
            return candidate

    raise FileNotFoundError(f"Tidak dapat menemukan berkas input di folder '{INPUT_DIR}' maupun '{BASE_DIR}'.")


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


def process_records_for_petugas(cleaner, petugas_name: str, raw_records: list, output_dir: Path):
    """Membersihkan dan mengekspor data untuk sekumpulan data mentah seorang petugas"""
    clean_name = petugas_name.replace(" ", "_").replace(".", "").replace(",", "")
    safe_name = "".join(c for c in clean_name if c.isalnum() or c == "_")

    cleaned_records, summary = cleaner.clean_records(raw_records)

    out_xlsx = output_dir / f"LK_Cleaning_Data_{safe_name}.xlsx"
    out_pdf = output_dir / f"Laporan_Cleaning_Data_{safe_name}.pdf"

    exporter = ExcelExporter(str(out_xlsx))
    exporter.export(cleaned_records, summary)

    reporter = PDFReporter(str(out_pdf))
    reporter.generate(cleaned_records, summary)

    pos_count = summary.total_usaha - summary.total_selisih_minus
    pos_pct = (pos_count / summary.total_usaha) * 100 if summary.total_usaha > 0 else 0
    warna_pct = (summary.total_warna_hijau / summary.total_usaha) * 100 if summary.total_usaha > 0 else 0
    rasio_pct = (summary.total_rasio_hijau / summary.total_usaha) * 100 if summary.total_usaha > 0 else 0
    unique_ratios = len(set(round(c.rasio_ntb_output, 4) for c in cleaned_records))

    print(f"  -> Usaha: {summary.total_usaha} | Warna Hijau: {warna_pct:.1f}% | Rasio NTB: {rasio_pct:.1f}% | Selisih >= 0: {pos_pct:.1f}% ({unique_ratios} rasio unik)")
    print(f"  -> Excel: {out_xlsx.name}")
    print(f"  -> PDF  : {out_pdf.name}")
    return summary


def process_petugas(cleaner: DataCleaner, petugas_name: str, output_dir: Path):
    """Menjalankan alur lengkap pembersihan dan pembuatan berkas untuk satu petugas spesifik"""
    print(f"\n[INFO] Membaca data untuk petugas: '{petugas_name}'...")
    raw_records = cleaner.load_records_by_petugas(petugas_name)

    if not raw_records:
        print(f"[PERINGATAN] Tidak ditemukan data untuk petugas '{petugas_name}'.")
        return None

    print(f"[INFO] Ditemukan {len(raw_records)} usaha. Menjalankan algoritma pembersihan data...")
    return process_records_for_petugas(cleaner, petugas_name, raw_records, output_dir)


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
        "--all", "--all-petugas",
        dest="all_petugas",
        action="store_true",
        help="Proses seluruh petugas (semua 27 petugas) sekaligus dalam satu kali eksekusi"
    )
    parser.add_argument(
        "--output-file",
        type=str,
        default="LK_Cleaning_Data_Gabungan.xlsx",
        help="Nama berkas Excel hasil gabungan seluruh petugas dalam 1 sheet (default: LK_Cleaning_Data_Gabungan.xlsx)"
    )
    parser.add_argument(
        "--split", "--split-petugas",
        dest="split_petugas",
        action="store_true",
        help="Selain berkas gabungan, juga buat berkas Excel dan PDF terpisah untuk masing-masing petugas"
    )
    parser.add_argument(
        "--input",
        type=str,
        help="Jalur khusus ke berkas Excel input (opsional)"
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=str(OUTPUT_DIR),
        help="Direktori penyimpanan hasil output Excel dan PDF (default: data/output)"
    )
    parser.add_argument(
        "--mode",
        type=str,
        choices=["lapangan", "standard"],
        default="lapangan",
        help="Mode pembersihan: 'lapangan' (sesuai lapangan, max 2x) atau 'standard'"
    )
    parser.add_argument(
        "--target-ntb",
        type=float,
        default=0.72,
        help="Target rasio NTB / Output (default: 0.72)"
    )
    parser.add_argument(
        "--rentang-bawah",
        type=float,
        default=0.64,
        help="Batas bawah rasio acuan NTB (default: 0.64)"
    )
    parser.add_argument(
        "--rentang-atas",
        type=float,
        default=0.78,
        help="Batas atas rasio acuan NTB (default: 0.78)"
    )
    parser.add_argument(
        "--min-surplus",
        type=float,
        default=0.0,
        help="Minimal surplus selisih NTB dalam Rupiah (default: 0.0, misal: 500000)"
    )
    parser.add_argument(
        "--max-multiplier",
        type=float,
        default=2.0,
        help="Batas maksimal kelipatan omzet dari nominal awal (default: 2.0, misal: 1.5)"
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

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.mode == "lapangan":
        print(
            f"[INFO] Menggunakan Mode 'Sesuai Lapangan' | Target NTB: {args.target_ntb:.2f} "
            f"| Rentang: [{args.rentang_bawah:.2f}, {args.rentang_atas:.2f}] "
            f"| Min Surplus: Rp {args.min_surplus:,.0f} | Max Multiplier: {args.max_multiplier:.1f}x"
        )
        cleaner = DataCleanerLapangan(
            str(input_file),
            target_ntb_ratio=args.target_ntb,
            min_ntb_ratio=args.rentang_bawah,
            max_ntb_ratio=args.rentang_atas,
            max_multiplier=args.max_multiplier,
            min_surplus=args.min_surplus,
        )
    else:
        print("[INFO] Menggunakan Mode 'Standard'")
        cleaner = DataCleaner(str(input_file))

    if args.all_petugas:
        print("\n[INFO] Opsi --all diaktifkan: Membaca seluruh data sheet LK sekaligus...")
        all_grouped = cleaner.load_all_records_grouped()
        total_petugas = len(all_grouped)
        total_usaha_semua = sum(len(v) for v in all_grouped.values())
        print(f"[INFO] Ditemukan {total_petugas} petugas dengan total {total_usaha_semua} unit usaha.")

        all_cleaned_records = []
        summaries = []

        for idx, (petugas, raw_records) in enumerate(sorted(all_grouped.items(), key=lambda x: x[0]), start=1):
            cleaned, summary = cleaner.clean_records(raw_records)
            summaries.append(summary)
            all_cleaned_records.extend(cleaned)

            if args.split_petugas:
                print(f"  [{idx}/{total_petugas}] Mengekspor berkas perorangan: '{petugas}'...")
                process_records_for_petugas(cleaner, petugas, raw_records, output_dir)

        # Urutkan seluruh baris data gabungan berdasarkan nomor baris asli dari sheet LK
        all_cleaned_records.sort(key=lambda c: c.row_idx)

        # Buat Ringkasan Eksekutif Terpadu
        cogs_ratios = [c.rasio_cogs for c in all_cleaned_records if c.rasio_cogs is not None]
        combined_summary = CleaningSummary(
            petugas=f"Semua Petugas ({total_petugas} Petugas)",
            total_usaha=len(all_cleaned_records),
            total_selesai=sum(1 for c in all_cleaned_records if c.sudah_diperbaiki),
            total_warna_hijau=sum(1 for c in all_cleaned_records if c.rasio_cogs is None or MIN_COGS_MARGIN <= c.rasio_cogs <= MAX_COGS_MARGIN),
            total_rasio_hijau=sum(1 for c in all_cleaned_records if c.status_rasio == "DALAM RENTANG ACUAN"),
            total_selisih_minus=sum(1 for c in all_cleaned_records if c.selisih_ntb < -0.01),
            total_ntb_awal=sum(c.ntb_awal for c in all_cleaned_records),
            total_ntb_adj=sum(c.ntb_adj for c in all_cleaned_records),
            total_selisih_ntb=sum(c.selisih_ntb for c in all_cleaned_records),
            min_cogs_ratio=min(cogs_ratios, default=0.0),
            max_cogs_ratio=max(cogs_ratios, default=0.0),
            min_ntb_ratio=min((c.rasio_ntb_output for c in all_cleaned_records), default=0.0),
            max_ntb_ratio=max((c.rasio_ntb_output for c in all_cleaned_records), default=0.0),
        )

        out_xlsx = output_dir / args.output_file
        pdf_name = args.output_file.replace(".xlsx", ".pdf") if args.output_file.endswith(".xlsx") else f"{args.output_file}.pdf"
        out_pdf = output_dir / f"Laporan_{pdf_name}"

        print(f"\n[INFO] Menyimpan seluruh data {total_petugas} petugas ke dalam 1 sheet ('LK') pada berkas gabungan:")
        print(f"  -> Excel Gabungan : {out_xlsx.resolve()}")
        exporter = ExcelExporter(str(out_xlsx))
        exporter.export(all_cleaned_records, combined_summary, sheet_name="LK", petugas_summaries=summaries)

        print(f"  -> PDF Rekapitulasi: {out_pdf.resolve()}")
        reporter = PDFReporter(str(out_pdf))
        reporter.generate(all_cleaned_records, combined_summary)

        print("\n" + "=" * 80)
        print("=== REKAPITULASI PROSES SELURUH PETUGAS ===")
        print("=" * 80)
        print(f"{'No':<3} | {'Nama Petugas':<40} | {'Usaha':<6} | {'Warna Hijau':<11} | {'Rasio Hijau':<11} | {'Bebas Minus'}")
        print("-" * 80)
        for i, s in enumerate(summaries, start=1):
            w_pct = (s.total_warna_hijau / s.total_usaha) * 100 if s.total_usaha > 0 else 0
            r_pct = (s.total_rasio_hijau / s.total_usaha) * 100 if s.total_usaha > 0 else 0
            minus_txt = "0 Minus (OK)" if s.total_selisih_minus == 0 else f"{s.total_selisih_minus} Minus"
            print(f"{i:<3} | {s.petugas:<40} | {s.total_usaha:<6} | {w_pct:>5.1f}%       | {r_pct:>5.1f}%       | {minus_txt}")
        print("-" * 80)
        print(f"Total Usaha Selesai Diproses: {total_usaha_semua} usaha untuk {total_petugas} petugas.")
        print(f"Seluruh 27 petugas tergabung dalam 1 sheet ('LK') pada: {out_xlsx.name}")
        print(f"Seluruh berkas Excel dan PDF tersimpan di: {output_dir.resolve()}")
    else:
        if not args.petugas:
            print("[INFO] Parameter --petugas tidak diberikan. Menjalankan default: 'Fadil Julianto'")
            petugas_target = "Fadil Julianto"
        else:
            petugas_target = args.petugas
        process_petugas(cleaner, petugas_target, output_dir)


if __name__ == "__main__":
    main()
