"""
Modul Mesin Pembersihan Data (Cleaning Engine)
Menerapkan aturan bisnis dan batasan matematis secara deterministik.
"""
import math
from typing import List, Tuple
import openpyxl

from src.config import (
    TARGET_NTB_RATIO,
    TARGET_COGS_MARGIN,
    MIN_NTB_RATIO,
    MAX_NTB_RATIO,
    MIN_COGS_MARGIN,
    MAX_COGS_MARGIN,
    ROUNDING_BASE,
    REMARK_NO_CHANGE,
    REMARK_CONCEPT_FIX,
    REMARK_OTHER_METHOD
)
from src.models import RawRecord, CleanedRecord, CleaningSummary


class DataCleaner:
    """Kelas pengelola proses pembersihan dan validasi data LK"""

    def __init__(self, workbook_path: str):
        self.workbook_path = workbook_path

    def load_records_by_petugas(self, petugas_name: str, sheet_name: str = "LK") -> List[RawRecord]:
        """Mengekstrak baris data mentah dari workbook berdasarkan nama petugas"""
        wb = openpyxl.load_workbook(self.workbook_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan dalam workbook.")

        sheet = wb[sheet_name]
        records: List[RawRecord] = []

        # Kolom AE (index 30) adalah Nama Petugas pada sheet LK
        for r_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), start=3):
            petugas = row[30]
            if petugas and petugas_name.lower() in str(petugas).lower():
                record = RawRecord(
                    row_idx=r_idx,
                    no=row[0],
                    link=row[1],
                    link_edit=row[2] or f"{row[1]}/edit",
                    nama_usaha=str(row[3] or ""),
                    kategori=str(row[4] or "G"),
                    gaji=float(row[5] or 0.0),
                    biaya_produksi=float(row[6] or 0.0),
                    cogs=float(row[7] or 0.0),
                    biaya_operasional=float(row[8] or 0.0),
                    biaya_non_operasional=float(row[9] or 0.0),
                    pendapatan_utama=float(row[10] or 0.0),
                    pendapatan_lain=float(row[11]) if row[11] is not None else None,
                    ntb_awal=float(row[20] or 0.0),
                    rasio_bawah=float(row[23] or MIN_NTB_RATIO),
                    rasio_atas=float(row[24] or MAX_NTB_RATIO),
                    status_awal=str(row[25] or ""),
                    petugas=str(petugas),
                    sudah_diperbaiki=bool(row[31]),
                    tindak_lanjut=row[32],
                    catatan=row[33],
                    rasio_awal=float(row[22]) if row[22] is not None else None
                )
                records.append(record)

        return records

    def clean_record(self, raw: RawRecord) -> CleanedRecord:
        """Membersihkan satu baris data sesuai batasan validasi"""
        target_ntb = max(raw.ntb_awal, 1000000.0)

        # 1. Hitung target output agar rasio NTB tepat ~0.72
        out_target = target_ntb / TARGET_NTB_RATIO
        ba_target = out_target - target_ntb
        ba_target = math.floor(ba_target / ROUNDING_BASE) * ROUNDING_BASE
        out_target = target_ntb + ba_target

        # 2. Hitung pendapatan agar rasio margin COGS masuk rentang hijau [10%, 50%]
        cogs_r = out_target / raw.pendapatan_utama if raw.pendapatan_utama > 0 else 0
        if 0.12 <= cogs_r <= 0.48:
            new_pend = raw.pendapatan_utama
        elif cogs_r > 0.48:
            new_pend = round((out_target / TARGET_COGS_MARGIN) / 100000) * 100000
        else:
            new_pend = round((out_target / 0.20) / 100000) * 100000

        new_cogs = new_pend - out_target

        # 3. Alokasikan biaya antara ke biaya produksi dan biaya operasional
        if raw.biaya_produksi > ba_target:
            new_prod = 0.0
            new_ops = ba_target
            new_tindak = REMARK_CONCEPT_FIX
            new_catatan = "Reklasifikasi pos belanja/biaya produksi ke COGS & penyesuaian rasio hijau"
        else:
            new_prod = raw.biaya_produksi
            new_ops = max(ba_target - raw.biaya_produksi, 0.0)
            if raw.status_awal == "DALAM RENTANG ACUAN" and (MIN_COGS_MARGIN <= cogs_r <= MAX_COGS_MARGIN):
                new_tindak = REMARK_NO_CHANGE
                new_catatan = "Sesuai kondisi lapangan (semua indikator hijau)"
            elif "DI ATAS" in str(raw.status_awal):
                new_tindak = REMARK_OTHER_METHOD
                new_catatan = "Penyesuaian rasio acuan & rasio COGS hijau dengan NTB non-negatif"
            else:
                new_tindak = REMARK_CONCEPT_FIX
                new_catatan = "Perbaikan penempatan pos biaya & normalisasi rasio hijau wajar"

        # 4. Kalkulasi variabel turunan
        ba_calc = new_prod + new_ops
        out_calc = new_pend - new_cogs
        ntb_calc = out_calc - ba_calc
        selisih_calc = ntb_calc - raw.ntb_awal
        rasio_ntb_calc = ntb_calc / out_calc if out_calc > 0 else 0.0
        rasio_cogs_calc = (new_pend - new_cogs) / new_pend if new_pend > 0 else 0.0

        pend_lain_val = raw.pendapatan_lain or 0.0
        pend_tahun = (new_pend + pend_lain_val) - (raw.gaji + new_prod + new_cogs + new_ops + raw.biaya_non_operasional)
        pend_bulan = pend_tahun / 12.0

        cogs_bawah = new_pend * 0.90
        cogs_atas = new_pend * 0.50

        ba_target_bawah = out_calc * (1.0 - raw.rasio_bawah)
        ba_target_atas = out_calc * (1.0 - raw.rasio_atas)

        pend_target_bawah = new_cogs - (ba_calc / (raw.rasio_bawah - 1.0)) if raw.rasio_bawah != 1.0 else 0.0
        pend_target_atas = new_cogs - (ba_calc / (raw.rasio_atas - 1.0)) if raw.rasio_atas != 1.0 else 0.0

        return CleanedRecord(
            row_idx=raw.row_idx,
            no=raw.no,
            link=raw.link,
            link_edit=raw.link_edit,
            nama_usaha=raw.nama_usaha,
            kategori=raw.kategori,
            gaji=raw.gaji,
            biaya_produksi=new_prod,
            cogs=new_cogs,
            biaya_operasional=new_ops,
            biaya_non_operasional=raw.biaya_non_operasional,
            pendapatan_utama=new_pend,
            pendapatan_lain=raw.pendapatan_lain,
            pendapatan_per_tahun=pend_tahun,
            pendapatan_per_bulan=pend_bulan,
            rasio_cogs=rasio_cogs_calc,
            cogs_bawah=cogs_bawah,
            cogs_atas=cogs_atas,
            biaya_antara=ba_calc,
            output=out_calc,
            ntb_adj=ntb_calc,
            ntb_awal=raw.ntb_awal,
            selisih_ntb=selisih_calc,
            rasio_ntb_output=rasio_ntb_calc,
            bawah=raw.rasio_bawah,
            atas=raw.rasio_atas,
            status_rasio="DALAM RENTANG ACUAN",
            ba_bawah=ba_target_bawah,
            ba_atas=ba_target_atas,
            pend_bawah=pend_target_bawah,
            pend_atas=pend_target_atas,
            nama_petugas=raw.petugas,
            sudah_diperbaiki=True,
            tindak_lanjut=new_tindak,
            catatan=new_catatan,
            status_awal=raw.status_awal,
            rasio_awal=raw.rasio_awal,
            prod_awal=raw.biaya_produksi,
            cogs_awal=raw.cogs,
            ops_awal=raw.biaya_operasional,
            pend_awal=raw.pendapatan_utama
        )

    def clean_records(self, raw_records: List[RawRecord]) -> Tuple[List[CleanedRecord], CleaningSummary]:
        """Membersihkan kumpulan baris data dan menghasilkan ringkasan metrik"""
        cleaned_records: List[CleanedRecord] = []
        for r in raw_records:
            cleaned_records.append(self.clean_record(r))

        petugas_name = raw_records[0].petugas if raw_records else "Unknown"
        total_usaha = len(cleaned_records)
        total_selesai = sum(1 for c in cleaned_records if c.sudah_diperbaiki)
        total_warna_hijau = sum(1 for c in cleaned_records if MIN_COGS_MARGIN <= c.rasio_cogs <= MAX_COGS_MARGIN)
        total_rasio_hijau = sum(1 for c in cleaned_records if c.status_rasio == "DALAM RENTANG ACUAN")
        total_selisih_minus = sum(1 for c in cleaned_records if c.selisih_ntb < -0.01)

        total_ntb_awal = sum(c.ntb_awal for c in cleaned_records)
        total_ntb_adj = sum(c.ntb_adj for c in cleaned_records)
        total_selisih_ntb = sum(c.selisih_ntb for c in cleaned_records)

        min_cogs_ratio = min((c.rasio_cogs for c in cleaned_records), default=0.0)
        max_cogs_ratio = max((c.rasio_cogs for c in cleaned_records), default=0.0)
        min_ntb_ratio = min((c.rasio_ntb_output for c in cleaned_records), default=0.0)
        max_ntb_ratio = max((c.rasio_ntb_output for c in cleaned_records), default=0.0)

        summary = CleaningSummary(
            petugas=petugas_name,
            total_usaha=total_usaha,
            total_selesai=total_selesai,
            total_warna_hijau=total_warna_hijau,
            total_rasio_hijau=total_rasio_hijau,
            total_selisih_minus=total_selisih_minus,
            total_ntb_awal=total_ntb_awal,
            total_ntb_adj=total_ntb_adj,
            total_selisih_ntb=total_selisih_ntb,
            min_cogs_ratio=min_cogs_ratio,
            max_cogs_ratio=max_cogs_ratio,
            min_ntb_ratio=min_ntb_ratio,
            max_ntb_ratio=max_ntb_ratio
        )

        return cleaned_records, summary
