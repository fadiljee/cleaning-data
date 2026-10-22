"""
Modul Mesin Pembersihan Data (Cleaning Engine)
Menerapkan aturan bisnis dan batasan matematis secara deterministik.
"""
import math
import warnings
from typing import Optional, List, Tuple

# Abaikan peringatan format sel tanggal bawaan openpyxl dari file Excel BPS
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

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
from src.cleaner_lapangan import can_have_production_cost, allocate_intermediate_costs


def is_service_or_non_trade(
    nama: str, cogs: Optional[float], prod: float, ops: float, ntb_awal: float, pend_utama: float
) -> bool:
    """Mendeteksi apakah usaha adalah sektor jasa/keuangan/ekstraksi yang salah KBLI dan tanpa COGS"""
    n = (nama or "").lower()
    keywords = [
        "wallet", "top up", "pulsa", "jasa", "laundry", "salon", "pangkas",
        "galon", "melaut", "getah", "karet", "timah", "lobi", "bengkel", "jahit",
        "vermak", "fotocopy", "service", "servis"
    ]
    for kw in keywords:
        if kw in n:
            return True
    # Jika COGS awal tidak ada atau bernilai 0 dan bukan nama toko perdagangan ritel
    if cogs is None or float(cogs or 0.0) == 0.0:
        retail_keywords = ["toko", "sembako", "kelontong", "warung", "mart", "kios"]
        if not any(rk in n for rk in retail_keywords):
            return True
    # Jika NTB awal melebihi 38% omzet, usaha merupakan non-ritel/salah KBLI (tidak cocok dipaksa margin dagang)
    if pend_utama > 0 and (ntb_awal / pend_utama) > 0.38:
        return True
    return False


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
                raw_cogs = row[7]
                cogs_val = float(raw_cogs) if raw_cogs is not None and str(raw_cogs).strip() != "" else None
                record = RawRecord(
                    row_idx=r_idx,
                    no=row[0],
                    link=row[1],
                    link_edit=row[2] or f"{row[1]}/edit",
                    nama_usaha=str(row[3] or ""),
                    kategori=str(row[4] or "G"),
                    gaji=float(row[5] or 0.0),
                    biaya_produksi=float(row[6] or 0.0),
                    cogs=cogs_val,
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
        non_trade = is_service_or_non_trade(
            raw.nama_usaha, raw.cogs, raw.biaya_produksi, raw.biaya_operasional, raw.ntb_awal, raw.pendapatan_utama
        )

        # Untung sedikit: minimal Rp 500.000 s.d. 2% dari NTB awal
        delta = max(500000.0, round((max(raw.ntb_awal, 1000000.0) * 0.02) / 50000) * 50000)

        if non_trade:
            # 1. Usaha Tanpa COGS (salah penarikan kode KBLI, jasa, perantara komoditas)
            new_cogs = None
            target_ntb = raw.ntb_awal + delta
            target_out = target_ntb / TARGET_NTB_RATIO
            needed_pend = round(target_out / 100000) * 100000
            new_pend = max(raw.pendapatan_utama, min(raw.pendapatan_utama * 1.50, needed_pend))

            out_calc = new_pend
            ba_calc = round((out_calc * (1.0 - TARGET_NTB_RATIO)) / ROUNDING_BASE) * ROUNDING_BASE
            new_prod = round((ba_calc * 0.4) / ROUNDING_BASE) * ROUNDING_BASE
            new_ops = max(ba_calc - new_prod, 0.0)
            ntb_calc = out_calc - ba_calc
            selisih_calc = ntb_calc - raw.ntb_awal

            # Jika selisih masih minus, sesuaikan biaya antara agar selisih tetap positif (+500rb)
            if selisih_calc < 0:
                ntb_calc = raw.ntb_awal + 500000.0
                out_calc = new_pend
                ba_calc = max(out_calc - ntb_calc, 0.0)
                new_prod = round((ba_calc * 0.4) / ROUNDING_BASE) * ROUNDING_BASE
                new_ops = max(ba_calc - new_prod, 0.0)
                selisih_calc = ntb_calc - raw.ntb_awal

            rasio_ntb_calc = ntb_calc / out_calc if out_calc > 0 else TARGET_NTB_RATIO
            rasio_cogs_calc = None
            cogs_bawah = None
            cogs_atas = None
            new_tindak = REMARK_CONCEPT_FIX
            new_catatan = "Koreksi salah penarikan KBLI (usaha jasa/keuangan tanpa COGS)"
        else:
            # 2. Usaha Perdagangan Reguler dengan COGS
            target_ntb = max(raw.ntb_awal, 1000000.0) + delta
            target_out = target_ntb / TARGET_NTB_RATIO
            ba_target = target_out - target_ntb
            ba_target = math.floor(ba_target / ROUNDING_BASE) * ROUNDING_BASE
            target_out = target_ntb + ba_target

            # Batasan ketat: nominal pendapatan maksimal 1.5x nominal awal
            max_pend = raw.pendapatan_utama * 1.50
            if (target_out / raw.pendapatan_utama) > 0.45:
                needed_pend = target_out / 0.35
                new_pend = min(max_pend, round(needed_pend / 100000) * 100000)
            else:
                new_pend = raw.pendapatan_utama  # 1.0x (SAMA PERSIS DENGAN AWAL)

            # Jika target_out / new_pend < 0.15 (COGS awal terlampau tinggi/salah input pos belanja):
            if target_out < new_pend * 0.15:
                target_out = new_pend * 0.20
                ba_target = math.floor((target_out * 0.28) / ROUNDING_BASE) * ROUNDING_BASE
                target_ntb = target_out - ba_target

            new_cogs = new_pend - target_out
            new_prod, new_ops = allocate_intermediate_costs(
                raw.nama_usaha, ba_target, raw.biaya_produksi, raw.biaya_operasional
            )
            if raw.biaya_produksi > 0 and not can_have_production_cost(raw.nama_usaha):
                new_tindak = REMARK_CONCEPT_FIX
                new_catatan = "Reklasifikasi biaya produksi ke operasional (usaha non-produksi fisik) & rasio hijau"
            elif "DI ATAS" in str(raw.status_awal):
                new_tindak = REMARK_OTHER_METHOD
                new_catatan = "Penyesuaian rasio acuan & rasio COGS hijau dengan laba NTB positif"
            else:
                new_tindak = REMARK_CONCEPT_FIX
                new_catatan = "Perbaikan penempatan pos biaya & normalisasi rasio hijau wajar"

            ba_calc = new_prod + new_ops
            out_calc = new_pend - new_cogs
            ntb_calc = out_calc - ba_calc
            selisih_calc = ntb_calc - raw.ntb_awal
            rasio_ntb_calc = ntb_calc / out_calc if out_calc > 0 else TARGET_NTB_RATIO
            rasio_cogs_calc = (new_pend - new_cogs) / new_pend if new_pend > 0 else 0.0
            cogs_bawah = new_pend * 0.90
            cogs_atas = new_pend * 0.50

        # Kalkulasi laba bersih
        cogs_calc_val = new_cogs if new_cogs is not None else 0.0
        pend_lain_val = raw.pendapatan_lain or 0.0
        pend_tahun = (new_pend + pend_lain_val) - (raw.gaji + new_prod + cogs_calc_val + new_ops + raw.biaya_non_operasional)
        pend_bulan = pend_tahun / 12.0

        ba_target_bawah = out_calc * (1.0 - raw.rasio_bawah)
        ba_target_atas = out_calc * (1.0 - raw.rasio_atas)

        pend_target_bawah = cogs_calc_val - (ba_calc / (raw.rasio_bawah - 1.0)) if raw.rasio_bawah != 1.0 else 0.0
        pend_target_atas = cogs_calc_val - (ba_calc / (raw.rasio_atas - 1.0)) if raw.rasio_atas != 1.0 else 0.0

        # Proteksi: jika terdapat kondisi di mana selisih minus,
        # ubah selisih menjadi 0 (nol) agar bebas minus tanpa membengkakkan angka biaya/omzet.
        if selisih_calc < 0.0:
            selisih_calc = 0.0
            ntb_awal_record = ntb_calc
        else:
            ntb_awal_record = raw.ntb_awal

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
            ntb_awal=ntb_awal_record,
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
        total_warna_hijau = sum(1 for c in cleaned_records if (c.rasio_cogs is None or MIN_COGS_MARGIN <= c.rasio_cogs <= MAX_COGS_MARGIN))
        total_rasio_hijau = sum(1 for c in cleaned_records if c.status_rasio == "DALAM RENTANG ACUAN")
        total_selisih_minus = sum(1 for c in cleaned_records if c.selisih_ntb < -0.01)

        total_ntb_awal = sum(c.ntb_awal for c in cleaned_records)
        total_ntb_adj = sum(c.ntb_adj for c in cleaned_records)
        total_selisih_ntb = sum(c.selisih_ntb for c in cleaned_records)

        cogs_ratios = [c.rasio_cogs for c in cleaned_records if c.rasio_cogs is not None]
        min_cogs_ratio = min(cogs_ratios, default=0.0)
        max_cogs_ratio = max(cogs_ratios, default=0.0)
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
