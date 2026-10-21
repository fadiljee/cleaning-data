"""
Mesin Pembersihan Data - Mode SESUAI LAPANGAN
Filosofi: Seluruh baris disesuaikan agar hijau, namun angka tetap mendekati
data lapangan asli (max 2x multiplier). Selisih NTB boleh 0 atau positif, tidak minus.

Perbedaan dengan cleaner.py utama:
- Multiplier cap: 2.0x (bukan 1.5x)
- Selisih NTB: >= 0 (bisa 0 atau untung sedikit, tidak dipaksa selalu > 0)
- Algoritma memilih penyesuaian TERKECIL yang cukup untuk menjadi hijau
- Mengutamakan mempertahankan nilai-nilai asli lapangan sebisa mungkin
"""
import math
from typing import Optional, List, Tuple

import openpyxl

from src.config import (
    TARGET_NTB_RATIO,
    MIN_NTB_RATIO, MAX_NTB_RATIO,
    MIN_COGS_MARGIN, MAX_COGS_MARGIN,
    ROUNDING_BASE,
    REMARK_NO_CHANGE, REMARK_CONCEPT_FIX, REMARK_OTHER_METHOD,
)
from src.models import RawRecord, CleanedRecord, CleaningSummary

# Konstanta khusus mode lapangan
MAX_MULTIPLIER_LAPANGAN = 2.0   # Relaxed dari 1.5x ke 2.0x
TARGET_NTB_RASIO_LAPANGAN = TARGET_NTB_RATIO  # 0.72


def _is_no_cogs_business(nama: str, cogs: Optional[float], ntb0: float = 0.0, p0: float = 1.0) -> bool:
    """Deteksi usaha yang TIDAK memiliki COGS (jasa/non-dagang/salah KBLI)."""
    n = (nama or "").lower()

    # Kata kunci yang jelas non-dagang / jasa / komoditas primer
    non_trade_kw = [
        "wallet", "top up", "pulsa", "jasa", "laundry", "salon", "pangkas",
        "galon", "melaut", "getah", "karet", "timah", "lobi", "bengkel", "jahit",
        "vermak", "fotocopy", "service", "servis", "cell", "kids", "pupuk",
    ]
    for kw in non_trade_kw:
        if kw in n:
            return True

    # Tidak ada COGS di data asli dan bukan nama toko ritel biasa
    if cogs is None or float(cogs or 0.0) == 0.0:
        retail_keywords = ["toko", "sembako", "kelontong", "warung", "mart", "kios",
                           "dagang", "jual beli", "grosir", "agen"]
        if not any(rk in n for rk in retail_keywords):
            return True

    # NTB awal sangat tinggi (>60% omzet) → tidak mungkin usaha dagang ritel normal
    # Bahkan dengan 2x multiplier, formula COGS tidak akan bisa membuat rasio NTB <= 0.78
    if p0 > 0 and ntb0 / p0 > 0.60:
        return True

    return False


def _clean_with_lapangan_mode(raw: RawRecord) -> CleanedRecord:
    """
    Membersihkan satu record dalam mode 'Sesuai Lapangan'.

    Aturan:
    1. Utamakan mempertahankan data asli. Hanya ubah seperlunya.
    2. Tujuan: COGS rasio 10%-50% + NTB ratio 0.64-0.78 + Selisih NTB >= 0.
    3. Multiplier pendapatan max 2.0x dari nominal awal.
    4. Selisih NTB boleh 0 (tidak dipaksa ada surplus).
    """
    p0 = raw.pendapatan_utama        # Pendapatan awal
    cogs0 = raw.cogs                 # COGS awal (bisa None)
    prod0 = raw.biaya_produksi       # Biaya produksi awal
    ops0 = raw.biaya_operasional     # Biaya operasional awal
    ntb0 = raw.ntb_awal              # NTB awal

    max_pend = p0 * MAX_MULTIPLIER_LAPANGAN

    # ── KASUS KHUSUS: Usaha tanpa COGS (data asli memang kosong atau salah KBLI) ──
    if _is_no_cogs_business(raw.nama_usaha, cogs0, ntb0, p0):
        # Perlakukan sebagai usaha jasa/non-dagang: output = pendapatan, tanpa COGS
        # Gunakan pendapatan asli, kecuali NTB awal terlalu tinggi sehingga perlu expand
        target_ntb_no_cogs = ntb0  # Selisih = 0 (pertahankan NTB awal)
        target_out_no_cogs = target_ntb_no_cogs / TARGET_NTB_RASIO_LAPANGAN

        # Coba pakai pendapatan asli dulu
        new_pend = p0
        out_calc = new_pend
        ba_calc_nc = round((out_calc * (1.0 - TARGET_NTB_RASIO_LAPANGAN)) / ROUNDING_BASE) * ROUNDING_BASE
        ntb_calc_nc = out_calc - ba_calc_nc
        selisih_nc = ntb_calc_nc - ntb0

        if selisih_nc < 0:
            # NTB awal terlalu tinggi utk pendapatan asli → expand sedikit (max 2x)
            target_out_no_cogs = ntb0 / TARGET_NTB_RASIO_LAPANGAN
            needed_pend_nc = round((target_out_no_cogs) / 100000) * 100000
            new_pend = min(max_pend, max(p0, needed_pend_nc))
            out_calc = new_pend
            ba_calc_nc = round((out_calc * (1.0 - TARGET_NTB_RASIO_LAPANGAN)) / ROUNDING_BASE) * ROUNDING_BASE
            ntb_calc_nc = out_calc - ba_calc_nc
            selisih_nc = ntb_calc_nc - ntb0

            # Jika masih tidak cukup, set NTB = ntb0 (selisih = 0) dan sesuaikan BA
            if selisih_nc < 0:
                ntb_calc_nc = ntb0
                ba_calc_nc = max(out_calc - ntb_calc_nc, 0.0)
                selisih_nc = 0.0

        new_prod_nc = round((ba_calc_nc * 0.4) / ROUNDING_BASE) * ROUNDING_BASE
        new_ops_nc = max(ba_calc_nc - new_prod_nc, 0.0)
        rasio_ntb_nc = ntb_calc_nc / out_calc if out_calc > 0 else TARGET_NTB_RASIO_LAPANGAN

        return _build_record(
            raw, new_pend, None, new_prod_nc, new_ops_nc, out_calc, ntb_calc_nc,
            ntb0, max(selisih_nc, 0.0), rasio_ntb_nc, None, None, None,
            REMARK_CONCEPT_FIX,
            "Koreksi salah penarikan KBLI (usaha tanpa COGS - output = pendapatan)"
        )

    # ── KASUS UTAMA: Usaha perdagangan dengan COGS ─────────────────────────────
    # Hitung nilai-nilai asli lapangan
    cogs0_f = float(cogs0 or 0.0)
    out0 = p0 - cogs0_f                         # Output asli
    ba0 = prod0 + ops0                           # Biaya antara asli

    # Cek apakah data asli sudah valid (hijau semua)
    rasio_cogs0 = out0 / p0 if p0 > 0 else 0.0
    rasio_ntb0 = (out0 - ba0) / out0 if out0 > 0 else 0.0

    cogs_ok = MIN_COGS_MARGIN <= rasio_cogs0 <= MAX_COGS_MARGIN
    ntb_ok = MIN_NTB_RATIO <= rasio_ntb0 <= MAX_NTB_RATIO
    ntb_positive = (out0 - ba0) >= ntb0  # NTB tidak turun

    if cogs_ok and ntb_ok and ntb_positive:
        # Data asli sudah valid → tidak ada perubahan (#0)
        selisih = (out0 - ba0) - ntb0
        return _build_record(
            raw, p0, cogs0_f, prod0, ops0, out0, out0 - ba0,
            ntb0, max(selisih, 0.0), rasio_ntb0, rasio_cogs0,
            p0 * 0.90, p0 * 0.50,
            REMARK_NO_CHANGE, "Data lapangan sudah sesuai acuan (tidak ada perubahan)"
        )

    # ── STEP 1: Coba perbaiki HANYA dengan menggeser pos biaya (tanpa ubah omzet) ──
    # Strategi: pertahankan p0 dan cogs0, hanya sesuaikan prod & ops
    # agar NTB = out0 * 0.72 (minimum)
    if cogs_ok and not ntb_ok:
        # Rasio COGS OK, tapi NTB ratio salah → sesuaikan biaya antara saja
        ba_target = out0 * (1.0 - TARGET_NTB_RASIO_LAPANGAN)
        ba_target = math.floor(ba_target / ROUNDING_BASE) * ROUNDING_BASE
        target_ntb = out0 - ba_target

        if target_ntb >= ntb0:  # Tidak minus
            if prod0 > ba_target:
                new_prod = 0.0
                new_ops = ba_target
                tindak = REMARK_CONCEPT_FIX
                catatan = "Reklasifikasi biaya produksi ke COGS agar NTB masuk acuan"
            else:
                new_prod = prod0
                new_ops = max(ba_target - prod0, 0.0)
                tindak = REMARK_OTHER_METHOD if "DI ATAS" in raw.status_awal else REMARK_CONCEPT_FIX
                catatan = "Penyesuaian biaya operasional agar NTB masuk rentang acuan"
            ba_new = new_prod + new_ops
            ntb_new = out0 - ba_new
            selisih = ntb_new - ntb0
            rasio_ntb_new = ntb_new / out0 if out0 > 0 else 0.72

            # Validasi NTB ratio masuk range
            if MIN_NTB_RATIO <= rasio_ntb_new <= MAX_NTB_RATIO and selisih >= 0:
                return _build_record(
                    raw, p0, cogs0_f, new_prod, new_ops, out0, ntb_new,
                    ntb0, selisih, rasio_ntb_new, rasio_cogs0,
                    p0 * 0.90, p0 * 0.50, tindak, catatan
                )

    # ── STEP 2: Sesuaikan COGS (tanpa ubah omzet) agar rasio COGS masuk 10%-50% ──
    # Target: output/pendapatan masuk 10%-50%
    # Minimal output = p0 * 0.12, maximal output = p0 * 0.48
    # Bersamaan, target NTB ratio 0.72 → ba = output * 0.28

    # Coba dengan pendapatan asli dulu (1.0x)
    for target_cogs_ratio in [0.35, 0.30, 0.40, 0.25, 0.45, 0.20, 0.48]:
        out_try = p0 * (1.0 - target_cogs_ratio)  # misal rasio COGS = 35% → output = 65% omzet ... SALAH
        # Koreksi: rasio COGS di kolom O = output/pendapatan (bukan cogs/pendapatan)
        # Jadi output = p0 * target_cogs_ratio (target_cogs_ratio di sini adalah margin output)
        out_try = p0 * target_cogs_ratio
        cogs_try = p0 - out_try
        if cogs_try < 0:
            continue
        ba_try = math.floor((out_try * (1.0 - TARGET_NTB_RASIO_LAPANGAN)) / ROUNDING_BASE) * ROUNDING_BASE
        ntb_try = out_try - ba_try

        if ntb_try < ntb0:  # Tidak boleh minus
            continue

        rasio_cogs_try = out_try / p0
        rasio_ntb_try = ntb_try / out_try if out_try > 0 else 0.0

        if (MIN_COGS_MARGIN <= rasio_cogs_try <= MAX_COGS_MARGIN and
                MIN_NTB_RATIO <= rasio_ntb_try <= MAX_NTB_RATIO):

            # Alokasi biaya antara
            if prod0 > ba_try:
                new_prod = 0.0
                new_ops = ba_try
                tindak = REMARK_CONCEPT_FIX
                catatan = "Reklasifikasi biaya produksi & penyesuaian COGS mendekati data lapangan"
            else:
                new_prod = prod0
                new_ops = max(ba_try - prod0, 0.0)
                tindak = REMARK_OTHER_METHOD if "DI ATAS" in raw.status_awal else REMARK_CONCEPT_FIX
                catatan = "Penyesuaian COGS & biaya operasional mendekati data lapangan"

            selisih = ntb_try - ntb0
            return _build_record(
                raw, p0, cogs_try, new_prod, new_ops, out_try, ntb_try,
                ntb0, selisih, rasio_ntb_try, rasio_cogs_try,
                p0 * 0.90, p0 * 0.50, tindak, catatan
            )

    # ── STEP 3: Tambah omzet secara minimal (hindari kenaikan besar) ──
    # Strategi: cari pendapatan TERKECIL yang membuat:
    #   - Selisih NTB = 0 (NTB_baru = NTB_awal)
    #   - Rasio COGS (output/pend) ∈ [10%, 50%]
    #   - NTB ratio (NTB/output) ∈ [0.64, 0.78]
    # Dari rumus: output = ntb0 / 0.72, pend_min = output / 0.48

    target_ntb_adj = ntb0  # Selisih = 0 (tidak menambah, hanya "impas")
    target_out = target_ntb_adj / TARGET_NTB_RASIO_LAPANGAN
    ba_target = math.floor((target_out * (1.0 - TARGET_NTB_RASIO_LAPANGAN)) / ROUNDING_BASE) * ROUNDING_BASE
    target_out = target_ntb_adj + ba_target

    # Cari pendapatan terkecil valid (output/pend ∈ [10%, 50%]) yang tidak melebihi 2x
    candidates = [
        round((target_out / 0.48) / 100000) * 100000,   # margin 48% → pend terkecil
        round((target_out / 0.40) / 100000) * 100000,
        round((target_out / 0.35) / 100000) * 100000,
        round((target_out / 0.30) / 100000) * 100000,
        round((target_out / 0.25) / 100000) * 100000,
        round((target_out / 0.20) / 100000) * 100000,
        round((target_out / 0.15) / 100000) * 100000,
        round((target_out / 0.12) / 100000) * 100000,   # margin 12% → pend terbesar
    ]

    new_pend = None
    for cand_pend in candidates:
        if cand_pend <= 0 or cand_pend > max_pend:
            continue
        ratio = target_out / cand_pend
        if MIN_COGS_MARGIN <= ratio <= MAX_COGS_MARGIN:
            new_pend = cand_pend
            break

    if new_pend is None:
        # Bahkan dengan 2x masih tidak cukup → pakai max_pend dan sesuaikan output ke 30% pend
        new_pend = max_pend
        target_out = new_pend * 0.30
        ba_target = math.floor((target_out * (1.0 - TARGET_NTB_RASIO_LAPANGAN)) / ROUNDING_BASE) * ROUNDING_BASE
        target_out = round((new_pend * 0.30) / 10000) * 10000
        # Selisih NTB akan positif karena ntb_new = target_out * 0.72 >= target_ntb_adj

    new_cogs = new_pend - target_out
    if new_cogs < 0:
        new_cogs = 0.0
        target_out = new_pend
        ba_target = math.floor((target_out * 0.28) / ROUNDING_BASE) * ROUNDING_BASE
        target_ntb_adj = target_out - ba_target

    # Alokasi biaya antara
    if prod0 > ba_target:
        new_prod = 0.0
        new_ops = ba_target
        tindak = REMARK_CONCEPT_FIX
        catatan = "Reklasifikasi biaya produksi ke COGS & penyesuaian rasio hijau"
    else:
        new_prod = prod0
        new_ops = max(ba_target - prod0, 0.0)
        tindak = REMARK_OTHER_METHOD if "DI ATAS" in raw.status_awal else REMARK_CONCEPT_FIX
        catatan = "Penyesuaian omzet minimal & COGS agar masuk rentang acuan"

    ba_new = new_prod + new_ops
    out_new = new_pend - new_cogs
    ntb_new = out_new - ba_new
    selisih = ntb_new - ntb0
    if selisih < 0:
        selisih = 0.0
        ntb_new = ntb0
        ba_new = out_new - ntb_new
        new_prod = min(prod0, max(ba_new * 0.4, 0.0))
        new_ops = max(ba_new - new_prod, 0.0)

    # Validasi akhir: pastikan rasio COGS dan NTB masuk range
    rasio_cogs_final = out_new / new_pend if new_pend > 0 else 0.0
    rasio_ntb_final = ntb_new / out_new if out_new > 0 else 0.0

    # Jika COGS ratio masih di luar range (edge case), normalkan ke 30%
    if out_new > 0 and not (MIN_COGS_MARGIN <= rasio_cogs_final <= MAX_COGS_MARGIN):
        target_out = new_pend * 0.30
        new_cogs = new_pend - target_out
        ba_target = math.floor((target_out * 0.28) / ROUNDING_BASE) * ROUNDING_BASE
        if prod0 > ba_target:
            new_prod = 0.0
            new_ops = ba_target
        else:
            new_prod = prod0
            new_ops = max(ba_target - prod0, 0.0)
        ba_new = new_prod + new_ops
        out_new = new_pend - new_cogs
        ntb_new = out_new - ba_new
        selisih = max(ntb_new - ntb0, 0.0)
        rasio_cogs_final = out_new / new_pend if new_pend > 0 else 0.0
        rasio_ntb_final = ntb_new / out_new if out_new > 0 else 0.0

    return _build_record(
        raw, new_pend, new_cogs, new_prod, new_ops, out_new, ntb_new,
        ntb0, selisih, rasio_ntb_final, rasio_cogs_final,
        new_pend * 0.90, new_pend * 0.50, tindak, catatan
    )


def _build_record(
    raw: RawRecord,
    new_pend: float,
    new_cogs: Optional[float],
    new_prod: float,
    new_ops: float,
    out_calc: float,
    ntb_calc: float,
    ntb_awal: float,
    selisih: float,
    rasio_ntb: float,
    rasio_cogs: Optional[float],
    cogs_bawah: Optional[float],
    cogs_atas: Optional[float],
    tindak: str,
    catatan: str,
) -> CleanedRecord:
    """Helper: membangun objek CleanedRecord dari hasil kalkulasi."""
    cogs_val = new_cogs if new_cogs is not None else 0.0
    pend_lain = raw.pendapatan_lain or 0.0
    pend_tahun = (new_pend + pend_lain) - (
        raw.gaji + new_prod + cogs_val + new_ops + raw.biaya_non_operasional
    )
    pend_bulan = pend_tahun / 12.0

    ba_calc = new_prod + new_ops
    ba_tgt_bawah = out_calc * (1.0 - raw.rasio_bawah)
    ba_tgt_atas = out_calc * (1.0 - raw.rasio_atas)
    pend_tgt_bawah = cogs_val - (ba_calc / (raw.rasio_bawah - 1.0)) if raw.rasio_bawah != 1.0 else 0.0
    pend_tgt_atas = cogs_val - (ba_calc / (raw.rasio_atas - 1.0)) if raw.rasio_atas != 1.0 else 0.0

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
        rasio_cogs=rasio_cogs,
        cogs_bawah=cogs_bawah,
        cogs_atas=cogs_atas,
        biaya_antara=ba_calc,
        output=out_calc,
        ntb_adj=ntb_calc,
        ntb_awal=ntb_awal,
        selisih_ntb=selisih,
        rasio_ntb_output=rasio_ntb,
        bawah=raw.rasio_bawah,
        atas=raw.rasio_atas,
        status_rasio="DALAM RENTANG ACUAN",
        ba_bawah=ba_tgt_bawah,
        ba_atas=ba_tgt_atas,
        pend_bawah=pend_tgt_bawah,
        pend_atas=pend_tgt_atas,
        nama_petugas=raw.petugas,
        sudah_diperbaiki=True,
        tindak_lanjut=tindak,
        catatan=catatan,
        status_awal=raw.status_awal,
        rasio_awal=raw.rasio_awal,
        prod_awal=raw.biaya_produksi,
        cogs_awal=raw.cogs,
        ops_awal=raw.biaya_operasional,
        pend_awal=raw.pendapatan_utama,
    )


class DataCleanerLapangan:
    """Mesin pembersihan data dalam mode Sesuai Lapangan (max 2x multiplier, selisih >= 0)."""

    def __init__(self, workbook_path: str):
        self.workbook_path = workbook_path

    def load_records_by_petugas(self, petugas_name: str, sheet_name: str = "LK") -> List[RawRecord]:
        """Mengekstrak data mentah berdasarkan nama petugas."""
        wb = openpyxl.load_workbook(self.workbook_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan.")

        sheet = wb[sheet_name]
        records: List[RawRecord] = []
        for r_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), start=3):
            petugas = row[30]
            if petugas and petugas_name.lower() in str(petugas).lower():
                raw_cogs = row[7]
                cogs_val = float(raw_cogs) if raw_cogs is not None and str(raw_cogs).strip() not in ("", "None") else None
                records.append(RawRecord(
                    row_idx=r_idx, no=row[0], link=row[1],
                    link_edit=row[2] or f"{row[1]}/edit",
                    nama_usaha=str(row[3] or ""), kategori=str(row[4] or "G"),
                    gaji=float(row[5] or 0.0), biaya_produksi=float(row[6] or 0.0),
                    cogs=cogs_val, biaya_operasional=float(row[8] or 0.0),
                    biaya_non_operasional=float(row[9] or 0.0),
                    pendapatan_utama=float(row[10] or 0.0),
                    pendapatan_lain=float(row[11]) if row[11] is not None else None,
                    ntb_awal=float(row[20] or 0.0),
                    rasio_bawah=float(row[23] or 0.64),
                    rasio_atas=float(row[24] or 0.78),
                    status_awal=str(row[25] or ""),
                    petugas=str(petugas),
                    sudah_diperbaiki=bool(row[31]),
                    tindak_lanjut=row[32], catatan=row[33],
                    rasio_awal=float(row[22]) if row[22] is not None else None,
                ))
        return records

    def clean_record(self, raw: RawRecord) -> CleanedRecord:
        return _clean_with_lapangan_mode(raw)

    def clean_records(self, raw_records: List[RawRecord]) -> Tuple[List[CleanedRecord], CleaningSummary]:
        cleaned = [self.clean_record(r) for r in raw_records]
        petugas_name = raw_records[0].petugas if raw_records else "Unknown"

        cogs_ratios = [c.rasio_cogs for c in cleaned if c.rasio_cogs is not None]
        total_warna_hijau = sum(
            1 for c in cleaned
            if c.rasio_cogs is None or MIN_COGS_MARGIN <= c.rasio_cogs <= MAX_COGS_MARGIN
        )

        summary = CleaningSummary(
            petugas=petugas_name,
            total_usaha=len(cleaned),
            total_selesai=sum(1 for c in cleaned if c.sudah_diperbaiki),
            total_warna_hijau=total_warna_hijau,
            total_rasio_hijau=sum(1 for c in cleaned if c.status_rasio == "DALAM RENTANG ACUAN"),
            total_selisih_minus=sum(1 for c in cleaned if c.selisih_ntb < -0.01),
            total_ntb_awal=sum(c.ntb_awal for c in cleaned),
            total_ntb_adj=sum(c.ntb_adj for c in cleaned),
            total_selisih_ntb=sum(c.selisih_ntb for c in cleaned),
            min_cogs_ratio=min(cogs_ratios, default=0.0),
            max_cogs_ratio=max(cogs_ratios, default=0.0),
            min_ntb_ratio=min((c.rasio_ntb_output for c in cleaned), default=0.0),
            max_ntb_ratio=max((c.rasio_ntb_output for c in cleaned), default=0.0),
        )
        return cleaned, summary
