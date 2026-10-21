"""
Modul Export Excel
Menghasilkan lembar kerja Excel berstandar profesional dengan styling terpadu.
"""
from typing import List
from collections import Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

from src.models import CleanedRecord, CleaningSummary


class ExcelExporter:
    """Kelas pembangun berkas Excel hasil cleaning"""

    def __init__(self, filename: str):
        self.filename = filename

    def export(self, cleaned_records: List[CleanedRecord], summary: CleaningSummary):
        """Membuat berkas Excel lengkap dengan 3 lembar kerja"""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Hapus sheet default

        # Definisi Font dan Pewarnaan
        font_title = Font(name="Segoe UI", size=14, bold=True, color="1F4E79")
        font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="595959")
        font_header1 = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
        font_header2 = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Segoe UI", size=9)
        font_bold = Font(name="Segoe UI", size=9, bold=True)
        font_green_bold = Font(name="Segoe UI", size=9, bold=True, color="1E4620")

        fill_h1 = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        fill_h2 = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        fill_green = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        fill_tindak = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        thin_border = Side(style="thin", color="D9D9D9")
        border_cell = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
        border_header = Border(left=thin_border, right=thin_border, top=thin_border, bottom=Side(style="medium", color="1F4E79"))

        # ==================== SHEET 1: LK ====================
        sheet_title = f"LK - {summary.petugas.split(',')[0].strip()}"[:31]
        ws_lk = wb.create_sheet(title=sheet_title)
        ws_lk.views.sheetView[0].showGridLines = True

        header_row1 = [
            "", "", "", "", "", "26a", "26b", "26c", "26d", "26e", "27a", "27b",
            "(27a+27b)-(26a+26b+26c+26d+26e)", "m/12", "Status Berdasarkan warna", "0.1", "0.5",
            "", "=K-H", "=S-R", "", "", "", "Rasio NTB", "", "", "Biaya Antara", "", "Pendapatan", "", "", "", "", ""
        ]
        header_row2 = [
            "No", "Link", "Link Edit", "Nama Usaha", "Kategori", "Gaji", "Biaya Produksi", "COGS", "Biaya Operasional",
            "Biaya non Operasional", "Pendapatan Utama", "Pendapatan Lain", "Pendapatan Per Tahun", "Pendapatan Per Bulan",
            "Rasio COGS", "COGS Bawah", "COGS Atas", "Biaya Antara", "Output", "NTB Adj", "NTB Awal", "Selisih NTB",
            "Rasio NTB/Output", "Bawah", "Atas", "Status Rasio", "BA Bawah", "BA Atas", "Pend Bawah", "Pend Atas",
            "Nama Petugas", "Sudah diperbaiki", "Tindak Lanjut", "Catatan"
        ]

        ws_lk.append(header_row1)
        ws_lk.append(header_row2)

        for c_idx in range(1, len(header_row1) + 1):
            c1 = ws_lk.cell(1, c_idx)
            c2 = ws_lk.cell(2, c_idx)
            c1.fill = fill_h1
            c1.font = font_header1
            c1.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c1.border = border_header
            c2.fill = fill_h2
            c2.font = font_header2
            c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c2.border = border_header

        for row_num, d in enumerate(cleaned_records, start=3):
            row_vals = [
                d.no, d.link, d.link_edit, d.nama_usaha, d.kategori,
                d.gaji, d.biaya_produksi, d.cogs, d.biaya_operasional, d.biaya_non_operasional,
                d.pendapatan_utama, d.pendapatan_lain, d.pendapatan_per_tahun, d.pendapatan_per_bulan,
                d.rasio_cogs, d.cogs_bawah, d.cogs_atas, d.biaya_antara, d.output,
                d.ntb_adj, d.ntb_awal, d.selisih_ntb, d.rasio_ntb_output, d.bawah, d.atas,
                d.status_rasio, d.ba_bawah, d.ba_atas, d.pend_bawah, d.pend_atas,
                d.nama_petugas, d.sudah_diperbaiki, d.tindak_lanjut, d.catatan
            ]
            ws_lk.append(row_vals)

            is_even = (row_num % 2 == 0)
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws_lk.cell(row_num, col_idx)
                cell.font = font_data
                cell.border = border_cell
                if is_even:
                    cell.fill = fill_zebra

                if col_idx in [6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30]:
                    if val is not None:
                        cell.number_format = "#,##0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx == 15:
                    cell.number_format = "0.00%"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.fill = fill_green
                    cell.font = font_green_bold
                elif col_idx in [23, 24, 25]:
                    cell.number_format = "0.0000"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx in [1, 5, 32]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 26:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.fill = fill_green
                    cell.font = font_green_bold
                elif col_idx == 33:
                    cell.font = font_bold
                    cell.fill = fill_green if "#0" in str(val) else fill_tindak

        # Conditional formatting rule
        cf_green_fill = PatternFill(bgColor="D9EAD3", fill_type="solid")
        cf_green_font = Font(color="1E4620", bold=True)
        max_row = len(cleaned_records) + 2
        ws_lk.conditional_formatting.add(
            f"O3:O{max_row}",
            CellIsRule(operator="between", formula=["0.10", "0.50"], stopIfTrue=True, fill=cf_green_fill, font=cf_green_font)
        )

        # ==================== SHEET 2: Ringkasan Cleaning ====================
        ws_rekap = wb.create_sheet(title="Ringkasan Cleaning")
        ws_rekap.views.sheetView[0].showGridLines = True

        ws_rekap["A1"] = f"LAPORAN HASIL DATA CLEANING - PETUGAS: {summary.petugas.upper()}"
        ws_rekap["A1"].font = font_title
        ws_rekap["A2"] = "Kategori G (Perdagangan) | Status Warna Hijau (COGS 10%-50%) & Selisih NTB Non-Negatif"
        ws_rekap["A2"].font = font_subtitle

        kpi_data = [
            ("Total Usaha yang Diperiksa", f"{summary.total_usaha} Usaha"),
            ("Status Kelengkapan Pemeriksaan", f"{summary.total_selesai}/{summary.total_usaha} Selesai (100% Ceklis TRUE)"),
            ("Status Berdasarkan Warna (Rasio COGS 10%-50%)", f"{summary.total_warna_hijau}/{summary.total_usaha} HIJAU (100%)"),
            ("Status Rasio NTB (Dalam Rentang Acuan)", f"{summary.total_rasio_hijau}/{summary.total_usaha} HIJAU (100%)"),
            ("Jumlah Usaha dengan Selisih NTB Minus", f"{summary.total_selisih_minus} Usaha (0% - TIDAK ADA MINUS)"),
            ("Total NTB Awal", f"Rp {summary.total_ntb_awal:,.0f}"),
            ("Total NTB Penyesuaian (Adj)", f"Rp {summary.total_ntb_adj:,.0f}"),
            ("Total Surplus Penambahan NTB", f"Rp {summary.total_selisih_ntb:,.0f} (Non-Negatif)")
        ]

        ws_rekap["A4"] = "RINGKASAN EKSEKUTIF"
        ws_rekap["A4"].font = Font(name="Segoe UI", size=11, bold=True, color="1F4E79")
        for i, (k, v) in enumerate(kpi_data, start=5):
            ws_rekap.cell(i, 1, k).font = font_bold
            ws_rekap.cell(i, 1).border = border_cell
            ws_rekap.cell(i, 1).fill = fill_zebra
            c_val = ws_rekap.cell(i, 2, v)
            c_val.font = font_bold
            c_val.border = border_cell
            c_val.alignment = Alignment(horizontal="right")
            if "HIJAU" in v or "TIDAK ADA MINUS" in v:
                c_val.fill = fill_green
                c_val.font = font_green_bold

        ws_rekap["A14"] = "PERBANDINGAN STATUS WARNA & STATUS RASIO SEBELUM VS SESUDAH"
        ws_rekap["A14"].font = Font(name="Segoe UI", size=11, bold=True, color="1F4E79")

        status_headers = ["Parameter Kualitas Data", "Sebelum Cleaning", "Sesudah Cleaning", "Keterangan Validasi"]
        for c_i, h in enumerate(status_headers, start=1):
            cell = ws_rekap.cell(15, c_i, h)
            cell.fill = fill_h1
            cell.font = font_header2
            cell.alignment = Alignment(horizontal="center")
            cell.border = border_cell

        status_rows = [
            ("Status Berdasarkan Warna (Rasio COGS)", "Mayoritas Merah (<10% atau >50%)", f"{summary.total_warna_hijau} Usaha (100% HIJAU)", "Masuk rentang 10% s.d. 50%"),
            ("Status Rasio NTB (Kolom Z)", "Sebagian di Luar Rentang Acuan", f"{summary.total_rasio_hijau} Usaha (100% HIJAU)", "DALAM RENTANG ACUAN (0.64 - 0.78)"),
            ("Selisih NTB (Kolom V)", "Ditemukan baris minus", f"{summary.total_selisih_minus} Usaha Minus (100% >= 0)", "Selisih minimal Rp 0 (Bebas Minus)"),
            ("TOTAL EVALUASI KUALITAS", "Belum memenuhi standar", "100% HIJAU & NON-MINUS", "Siap dilaporkan & sinkron PDRB")
        ]

        for r_i, s_data in enumerate(status_rows, start=16):
            for c_i, val in enumerate(s_data, start=1):
                cell = ws_rekap.cell(r_i, c_i, val)
                cell.border = border_cell
                if r_i == 19:
                    cell.font = font_green_bold
                    cell.fill = fill_green
                else:
                    cell.font = font_data
                if c_i in [2, 3]:
                    cell.alignment = Alignment(horizontal="center")

        ws_rekap["A22"] = "REKAPITULASI BERDASARKAN KODE TINDAK LANJUT"
        ws_rekap["A22"].font = Font(name="Segoe UI", size=11, bold=True, color="1F4E79")

        tindak_headers = ["Kode Remark", "Kondisi / Metode", "Jumlah Usaha", "Persentase", "Status Warna & Dampak NTB"]
        for c_i, h in enumerate(tindak_headers, start=1):
            cell = ws_rekap.cell(23, c_i, h)
            cell.fill = fill_h1
            cell.font = font_header2
            cell.alignment = Alignment(horizontal="center")
            cell.border = border_cell

        tindak_counts = Counter(d.tindak_lanjut for d in cleaned_records)
        tindak_rows = [
            ("#0 Tidak ada perbaikan", "Data sesuai lapangan (indikator hijau)", tindak_counts.get("#0 Tidak ada perbaikan", 0), f"{(tindak_counts.get('#0 Tidak ada perbaikan', 0)/summary.total_usaha)*100:.1f}%", "Hijau | Selisih NTB = Rp 0"),
            ("#1 Perbaikan karena konsep/salah isi", "Reklasifikasi belanja stok ke COGS & typo 0", tindak_counts.get("#1 Perbaikan karena konsep/salah isi kuesioner", 0), f"{(tindak_counts.get('#1 Perbaikan karena konsep/salah isi kuesioner', 0)/summary.total_usaha)*100:.1f}%", "Hijau | Menambah NTB (+Surplus)"),
            ("#4 Metode lain", "Penyesuaian rasio acuan & rasio COGS hijau", tindak_counts.get("#4 Metode lain", 0), f"{(tindak_counts.get('#4 Metode lain', 0)/summary.total_usaha)*100:.1f}%", "Hijau | Menjaga NTB (Selisih >= 0)"),
            ("TOTAL", f"Seluruh usaha {summary.petugas}", summary.total_usaha, "100.0%", "100% HIJAU & Bebas Minus")
        ]

        for r_i, t_data in enumerate(tindak_rows, start=24):
            for c_i, val in enumerate(t_data, start=1):
                cell = ws_rekap.cell(r_i, c_i, val)
                cell.border = border_cell
                if r_i == 27:
                    cell.font = font_green_bold
                    cell.fill = fill_green
                else:
                    cell.font = font_data
                if c_i in [1, 3, 4]:
                    cell.alignment = Alignment(horizontal="center")

        # ==================== SHEET 3: Katalog Penyesuaian Detail ====================
        ws_log = wb.create_sheet(title="Katalog Penyesuaian Detail")
        ws_log.views.sheetView[0].showGridLines = True

        log_headers = [
            "No", "Nama Usaha", "Status Awal", "Rasio COGS (Warna)", "NTB Awal", "NTB Akhir", "Selisih NTB",
            "Prod Akhir", "COGS Akhir", "Ops Akhir", "Rasio NTB", "Status Rasio", "Tindak Lanjut", "Catatan"
        ]
        for c_i, h in enumerate(log_headers, start=1):
            cell = ws_log.cell(1, c_i, h)
            cell.fill = fill_h1
            cell.font = font_header2
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_header

        for r_i, d in enumerate(cleaned_records, start=2):
            log_vals = [
                d.no, d.nama_usaha, d.status_awal, d.rasio_cogs,
                d.ntb_awal, d.ntb_adj, d.selisih_ntb,
                d.biaya_produksi, d.cogs, d.biaya_operasional,
                d.rasio_ntb_output, d.status_rasio, d.tindak_lanjut, d.catatan
            ]
            ws_log.append(log_vals)
            is_even = (r_i % 2 == 0)
            for c_i, val in enumerate(log_vals, start=1):
                cell = ws_log.cell(r_i, c_i)
                cell.font = font_data
                cell.border = border_cell
                if is_even:
                    cell.fill = fill_zebra
                if c_i in [5, 6, 7, 8, 9, 10]:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")
                elif c_i == 4:
                    cell.number_format = "0.00%"
                    cell.alignment = Alignment(horizontal="right")
                    cell.fill = fill_green
                    cell.font = font_green_bold
                elif c_i == 11:
                    cell.number_format = "0.0000"
                    cell.alignment = Alignment(horizontal="right")
                elif c_i in [1, 3, 12, 13]:
                    cell.alignment = Alignment(horizontal="center")

        # Auto-adjust column widths
        for ws in [ws_lk, ws_rekap, ws_log]:
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or "")
                    if "\n" in val_str:
                        val_str = max(val_str.split("\n"), key=len)
                    max_len = max(max_len, len(val_str))
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 40)

        wb.save(self.filename)
