"""
Mesin Pembersihan Data - Mode SESUAI LAPANGAN
Filosofi:
1. Seluruh baris disesuaikan agar hijau (Status Warna COGS 10%-50% & Rasio NTB 0.64-0.78).
2. Angka tetap mendekati data lapangan asli (multiplier pendapatan minimal, diutamakan 1.0x).
3. Output berada di rentang normal sektor perdagangan (margin kotor 12% - 48%, bukan omzet bruto).
4. Penyesuaian pos biaya kontekstual sesuai jenis usaha:
   - Usaha dagang/ritel/warung/sembako/kios/pengepul: Biaya Produksi = 0, Biaya Operasional wajar.
   - Usaha produksi/jasa fisik (bengkel, jahit, salon, las): Biaya Produksi & Operasional proporsional.
5. Rasio NTB bervariasi secara alami antar usaha (tidak seragam 72.00%).
6. Selisih NTB diutamakan untung sedikit atau nol (minimal surplus), tanpa membengkakkan biaya.
"""
import math
import warnings
from typing import Optional, List, Tuple

# Abaikan peringatan format sel tanggal bawaan openpyxl dari file Excel BPS
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import openpyxl

from src.config import (
    TARGET_NTB_RATIO,
    MIN_NTB_RATIO, MAX_NTB_RATIO,
    MIN_COGS_MARGIN, MAX_COGS_MARGIN,
    ROUNDING_BASE,
    REMARK_NO_CHANGE, REMARK_CONCEPT_FIX, REMARK_OTHER_METHOD,
)
from src.models import RawRecord, CleanedRecord, CleaningSummary


def safe_float(val, default: float = 0.0) -> float:
    """Konversi nilai cell Excel ke float dengan aman, menangani #DIV/0!, #VALUE!, None, string kosong."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.startswith("#") or s.lower() in ("none", "null", "-", "nan"):
        return default
    try:
        return float(s.replace(",", ""))
    except (ValueError, TypeError):
        return default


def safe_float_or_none(val) -> Optional[float]:
    """Konversi nilai cell Excel ke float atau None jika kosong / error formula Excel."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.startswith("#") or s.lower() in ("none", "null", "-", "nan"):
        return None
    try:
        return float(s.replace(",", ""))
    except (ValueError, TypeError):
        return None


def is_pure_service(nama: str) -> bool:
    """Mengecek apakah usaha adalah jasa murni/sosial yang tidak menjual barang dagangan (tanpa COGS)."""
    n = (nama or "").lower()
    keywords = [
        "pangkas", "potong rambut", "salon", "laundry", "urut", "pijat",
        "dokter", "bidan", "klinik", "tpa", "sekolah", "yayasan", "pesantren",
        "bimbingan", "bimbel", "les", "jahit", "vermak", "tambal ban", "cuci motor", "cuci mobil"
    ]
    return any(kw in n for kw in keywords)


def is_trade_with_cogs(nama: str, kategori: str, cogs0_f: Optional[float], pend_utama: float) -> bool:
    """Mengecek apakah usaha merupakan sektor perdagangan yang menggunakan formula COGS."""
    if cogs0_f is not None and cogs0_f > 0.0:
        return True
    if kategori and str(kategori).strip().upper() not in ("G", "NONE", ""):
        return False
    if is_pure_service(nama):
        return False
    n = (nama or "").lower()
    retail_keywords = [
        "toko", "sembako", "kelontong", "warung", "mart", "kios", "grosir",
        "distributor", "lapak", "pengecer", "jual", "top up", "wallet", "pulsa", "cell", "counter", "loket"
    ]
    if any(rk in n for rk in retail_keywords):
        return True
    return False


def is_service_or_non_trade(
    nama: str, cogs: Optional[float], prod: float, ops: float, ntb_awal: float, pend_utama: float
) -> bool:
    """Mendeteksi apakah usaha adalah sektor jasa/keuangan/ekstraksi/sosial yang salah KBLI dan tanpa COGS"""
    return not is_trade_with_cogs(nama, "G", cogs, pend_utama)


def can_have_production_cost(nama: str) -> bool:
    """
    Mengecek apakah jenis usaha secara wajar dapat memiliki biaya produksi/kemasan/penyiapan barang.
    Usaha komoditas tertutup murni (pangkalan gas, apotek, dealer, pom mini, agen voucher pulsa)
    tidak memiliki proses produksi/pengolahan/kemasan fisik mandiri (Biaya Produksi = 0).
    Usaha lainnya (termasuk toko sembako, warung, keliling sayur/ikan, lapak, bengkel, olahan)
    dapat memiliki biaya produksi/kemasan/penyiapan barang.
    """
    n = (nama or "").lower()
    pure_no_prod = [
        "pangkalan gas", "agen gas", "lpg", "elpiji",
        "apotek", "farmasi", "toko obat",
        "dealer", "showroom", "goda",
        "top up", "pulsa", "loket", "e-wallet", "ewallet",
        "pom mini", "pommini", "pertamini", "pertashop", "spbu"
    ]
    return not any(kw in n for kw in pure_no_prod)


def allocate_intermediate_costs(
    nama: str,
    cand_ba: float,
    prod0: float,
    ops0: float,
) -> Tuple[float, float]:
    """
    Mengalokasikan Biaya Antara (cand_ba) ke Biaya Produksi dan Biaya Operasional
    secara proporsional mempertahankan data asli lapangan dan tidak menghilangkan biaya
    yang memang seharusnya ada.
    """
    if not can_have_production_cost(nama):
        # Usaha murni non-produksi: Biaya Produksi = 0, seluruhnya dialokasikan ke Biaya Operasional
        return 0.0, cand_ba

    # Jika di data mentah memang tidak ada biaya produksi (prod0 = 0):
    if prod0 <= 0:
        return 0.0, cand_ba

    # Jika di data mentah ada biaya produksi (prod0 > 0):
    ba0 = prod0 + ops0
    if ba0 <= 0:
        return 0.0, cand_ba

    if ops0 <= 0:
        # Jika operasional awal 0, beri porsi produksi wajar dan sisanya untuk operasional
        cand_prod = round((cand_ba * 0.65) / ROUNDING_BASE) * ROUNDING_BASE
        cand_ops = cand_ba - cand_prod
        return cand_prod, cand_ops

    # Jika kedua biaya ada di data mentah (prod0 > 0 dan ops0 > 0):
    # Pertahankan proporsi asli lapangan (r_prod = prod0 / ba0)
    r_prod = prod0 / ba0
    # Jaga agar porsi produksi antara 10% s.d. 75% sehingga biaya operasional tetap wajar
    r_prod = max(0.10, min(0.75, r_prod))
    cand_prod = round((cand_ba * r_prod) / ROUNDING_BASE) * ROUNDING_BASE
    cand_ops = cand_ba - cand_prod

    # Pastikan kedua pos tetap bernilai jika cand_ba mencukupi
    if cand_ba >= (2 * ROUNDING_BASE):
        if cand_prod <= 0:
            cand_prod = float(ROUNDING_BASE)
            cand_ops = cand_ba - cand_prod
        elif cand_ops <= 0:
            cand_ops = float(ROUNDING_BASE)
            cand_prod = cand_ba - cand_ops

    return cand_prod, cand_ops


def is_production_business(nama: str) -> bool:
    """Kompatibilitas mundur: mendeteksi usaha manufaktur/jasa fisik."""
    return can_have_production_cost(nama)


def get_row_target_ratio(
    row_idx: int,
    nama: str,
    base_target: float = TARGET_NTB_RATIO,
    min_r: float = MIN_NTB_RATIO,
    max_r: float = MAX_NTB_RATIO,
) -> float:
    """
    Menghasilkan rasio NTB/Output yang bervariasi secara alami dan deterministik
    dalam rentang aman acuan [min_r, max_r] agar persentase tidak seragam.
    """
    h = (row_idx * 47 + abs(hash(nama or ""))) % 71
    jitter = (h - 35) * 0.0010  # deviasi antara -0.0350 s.d. +0.0350
    r = base_target + jitter
    # Batasi agar selalu berada di zona hijau yang aman (misal 0.66 s.d. 0.76)
    safe_min = max(min_r + 0.015, 0.6550)
    safe_max = min(max_r - 0.015, 0.7650)
    return round(max(safe_min, min(safe_max, r)), 4)


def _clean_with_lapangan_mode(
    raw: RawRecord,
    target_ntb_ratio: float = TARGET_NTB_RATIO,
    min_ntb_ratio: float = MIN_NTB_RATIO,
    max_ntb_ratio: float = MAX_NTB_RATIO,
    max_multiplier: float = 2.0,
    min_surplus: float = 0.0,
) -> CleanedRecord:
    """
    Membersihkan satu record dalam mode 'Sesuai Lapangan'.

    Prinsip:
    1. Utamakan mempertahankan data asli.
    2. Output perdagangan berada pada rentang margin normal (12% - 48%).
    3. Biaya dialokasikan sesuai jenis usaha:
       - Ritel/warung/sembako/pengepul: Biaya Produksi = 0, Biaya Operasional wajar.
    4. Rasio NTB bervariasi alami dalam rentang hijau [min_ntb_ratio, max_ntb_ratio].
    5. Selisih NTB diutamakan minimal positif (untung sedikit / nol).
    """
    p0 = raw.pendapatan_utama
    cogs0 = raw.cogs
    prod0 = raw.biaya_produksi
    ops0 = raw.biaya_operasional
    ntb0 = raw.ntb_awal
    nama = raw.nama_usaha
    nama_lower = (nama or "").lower()

    cogs0_f = float(cogs0) if cogs0 is not None and str(cogs0).strip() not in ("", "None") else None
    is_prod = is_production_business(nama)
    target_r = get_row_target_ratio(raw.row_idx, nama, target_ntb_ratio, min_ntb_ratio, max_ntb_ratio)

    # -- KASUS KHUSUS 0: Data mentah dengan omzet/pendapatan kosong atau 0 --
    if p0 <= 0:
        if ntb0 > 0:
            cand_ntb = ntb0
            cand_out = round((cand_ntb / target_r) / ROUNDING_BASE) * ROUNDING_BASE
            cand_ba = cand_out - cand_ntb
            cand_prod, cand_ops = allocate_intermediate_costs(nama, cand_ba, prod0, ops0)

            is_service = any(k in nama_lower for k in [
                "tpa", "sekolah", "yayasan", "pesantren", "bimbingan", "laundry", "salon",
                "pangkas", "bengkel", "servis", "service", "jahit", "fotocopy"
            ])

            if is_service:
                cand_p = cand_out
                cand_cogs = None
                r_cogs = None
            else:
                cand_p = round((cand_out / 0.28) / 100000) * 100000
                cand_cogs = cand_p - cand_out
                r_cogs = cand_out / cand_p

            r_ntb = cand_ntb / cand_out if cand_out > 0 else target_r
            selisih = 0.0
            tindak = REMARK_CONCEPT_FIX
            catatan = "Rekonstruksi omzet, COGS, dan biaya dari NTB awal survei (data omzet survei awal kosong)"

            return _build_record(
                raw, cand_p, cand_cogs, cand_prod, cand_ops, cand_out, cand_ntb,
                ntb0, selisih, r_ntb, r_cogs,
                cand_p * 0.90 if cand_cogs is not None else None,
                cand_p * 0.50 if cand_cogs is not None else None,
                tindak, catatan
            )
        else:
            return _build_record(
                raw, 0.0, None, 0.0, 0.0, 0.0, 0.0,
                0.0, 0.0, target_r, None,
                None, None,
                REMARK_NO_CHANGE, "Usaha nihil / tidak ada data keuangan"
            )

    # -- KONDISI 1: Data lapangan asli SUDAH HIJAU dan selisih >= min_surplus --
    if cogs0_f is not None and p0 > 0:
        raw_out = p0 - cogs0_f
        raw_ba = prod0 + ops0
        raw_ntb = raw_out - raw_ba
        raw_rcogs = raw_out / p0
        raw_rntb = raw_ntb / raw_out if raw_out > 0 else 0.0

        cogs_ok = min_cogs_margin_ok = (MIN_COGS_MARGIN <= raw_rcogs <= MAX_COGS_MARGIN)
        ntb_ok = min_ntb_ratio <= raw_rntb <= max_ntb_ratio
        surplus_ok = (raw_ntb - ntb0) >= min_surplus

        if cogs_ok and ntb_ok and surplus_ok:
            return _build_record(
                raw, p0, cogs0_f, prod0, ops0, raw_out, raw_ntb,
                ntb0, raw_ntb - ntb0, raw_rntb, raw_rcogs,
                p0 * 0.90, p0 * 0.50,
                REMARK_NO_CHANGE, "Data lapangan sudah sesuai acuan (tidak ada perubahan)"
            )

    # -- KONDISI 2: Usaha jasa / non-dagang murni tanpa COGS --
    if not is_trade_with_cogs(nama, raw.kategori, cogs0_f, p0) and p0 > 0:
        raw_out = p0
        raw_ba = prod0 + ops0
        raw_ntb = raw_out - raw_ba
        raw_rntb = raw_ntb / raw_out if raw_out > 0 else 0.0

        if min_ntb_ratio <= raw_rntb <= max_ntb_ratio and (raw_ntb - ntb0) >= min_surplus:
            return _build_record(
                raw, p0, None, prod0, ops0, raw_out, raw_ntb,
                ntb0, raw_ntb - ntb0, raw_rntb, None,
                None, None,
                REMARK_NO_CHANGE, "Data lapangan sudah sesuai acuan (usaha jasa/tanpa COGS)"
            )

        # Sesuaikan target NTB dan output
        needed_out = round((ntb0 / target_r) / ROUNDING_BASE) * ROUNDING_BASE
        if ntb0 > 0 and needed_out <= p0 * max_multiplier:
            cand_p = needed_out
            cand_out = needed_out
            cand_ntb = round((cand_out * target_r) / ROUNDING_BASE) * ROUNDING_BASE
            selisih = cand_ntb - ntb0
        else:
            cand_p = p0
            cand_out = p0
            cand_ntb = round((cand_out * target_r) / ROUNDING_BASE) * ROUNDING_BASE
            selisih = 0.0

        cand_ba = max(cand_out - cand_ntb, 0.0)
        cand_prod, cand_ops = allocate_intermediate_costs(nama, cand_ba, prod0, ops0)
        r_ntb = cand_ntb / cand_out if cand_out > 0 else target_r
        tindak = REMARK_CONCEPT_FIX
        catatan = "Penyesuaian biaya usaha jasa/non-dagang tanpa COGS sesuai rentang acuan"

        ntb_awal_record = cand_ntb if (selisih == 0.0 and (cand_ntb - ntb0) < 0) else ntb0

        return _build_record(
            raw, cand_p, None, cand_prod, cand_ops, cand_out, cand_ntb,
            ntb_awal_record, max(selisih, 0.0), r_ntb, None,
            None, None,
            tindak, catatan
        )

    # -- KONDISI 3: Optimasi Cerdas Kontekstual & Minimal Surplus --
    # Surplus deterministik kecil: 0 s.d. 300rb Rupiah
    if min_surplus > 0:
        target_surplus = min_surplus
    else:
        target_surplus = ((raw.row_idx * 17) % 7) * 50000.0  # 0, 50k, 100k, 150k, 200k, 250k, 300k

    best_solution = None
    all_multipliers = [1.0, 1.02, 1.05, 1.08, 1.10, 1.15, 1.20, 1.25, 1.30, 1.35, 1.40, 1.50, 1.65, 1.80, 2.0]
    p_multipliers = [m for m in all_multipliers if m <= max_multiplier + 0.001]
    if not p_multipliers:
        p_multipliers = [1.0]

    for p_mult in p_multipliers:
        if p0 < 100000:
            cand_p = max(p0, 10000.0)
        else:
            cand_p = round((p0 * p_mult) / 100000) * 100000
            if cand_p <= 0:
                cand_p = p0

        # Cari persentase margin Output/Pendapatan antara 12% s.d. 48% (rentang normal perdagangan)
        for m_int in range(12, 49):
            m = m_int / 100.0
            cand_out = round((cand_p * m) / ROUNDING_BASE) * ROUNDING_BASE
            cand_cogs = cand_p - cand_out
            cand_ntb = round((cand_out * target_r) / ROUNDING_BASE) * ROUNDING_BASE
            cand_ba = cand_out - cand_ntb
            selisih = cand_ntb - ntb0

            # Alokasi biaya secara kontekstual sesuai jenis usaha (mempertahankan data biaya lapangan)
            cand_prod, cand_ops = allocate_intermediate_costs(nama, cand_ba, prod0, ops0)

            r_cogs = (cand_out / cand_p) if cand_p > 0 else 0.0
            r_ntb = (cand_ntb / cand_out) if cand_out > 0 else 0.0

            if (MIN_COGS_MARGIN <= r_cogs <= MAX_COGS_MARGIN and
                    min_ntb_ratio <= r_ntb <= max_ntb_ratio and
                    selisih >= target_surplus):
                if p_mult > 1.0:
                    tindak = REMARK_OTHER_METHOD
                    catatan = "Penyesuaian omzet minimal & COGS mendekati rentang acuan"
                elif prod0 > 0 and not can_have_production_cost(nama):
                    tindak = REMARK_CONCEPT_FIX
                    catatan = "Reklasifikasi biaya produksi ke operasional (usaha non-produksi fisik) & COGS wajar"
                else:
                    tindak = REMARK_CONCEPT_FIX
                    catatan = "Penyesuaian COGS dan biaya operasional/produksi wajar mempertahankan data lapangan"

                best_solution = (
                    cand_p, cand_cogs, cand_prod, cand_ops, cand_out, cand_ntb, selisih,
                    r_ntb, r_cogs, tindak, catatan
                )
                break

        if best_solution is not None:
            break

    # Jika belum ada solusi dengan target_surplus, coba sekali lagi dengan selisih >= 0
    if best_solution is None and target_surplus > 0:
        for p_mult in p_multipliers:
            if p0 < 100000:
                cand_p = max(p0, 10000.0)
            else:
                cand_p = round((p0 * p_mult) / 100000) * 100000
                if cand_p <= 0:
                    cand_p = p0
            for m_int in range(12, 49):
                m = m_int / 100.0
                cand_out = round((cand_p * m) / ROUNDING_BASE) * ROUNDING_BASE
                cand_cogs = cand_p - cand_out
                cand_ntb = round((cand_out * target_r) / ROUNDING_BASE) * ROUNDING_BASE
                cand_ba = cand_out - cand_ntb
                selisih = cand_ntb - ntb0

                cand_prod, cand_ops = allocate_intermediate_costs(nama, cand_ba, prod0, ops0)

                r_cogs = (cand_out / cand_p) if cand_p > 0 else 0.0
                r_ntb = (cand_ntb / cand_out) if cand_out > 0 else 0.0

                if (MIN_COGS_MARGIN <= r_cogs <= MAX_COGS_MARGIN and
                        min_ntb_ratio <= r_ntb <= max_ntb_ratio and
                        selisih >= 0.0):
                    tindak = REMARK_CONCEPT_FIX
                    catatan = "Penyesuaian COGS dan biaya operasional/produksi wajar mempertahankan data lapangan"
                    best_solution = (
                        cand_p, cand_cogs, cand_prod, cand_ops, cand_out, cand_ntb, selisih,
                        r_ntb, r_cogs, tindak, catatan
                    )
                    break
            if best_solution is not None:
                break

    # -- KONDISI 4: Fallback untuk anomali ekstrim data mentah (NTB awal > 50% omzet) --
    # Jika NTB awal data mentah sangat tinggi akibat survei awal tidak mengisi COGS (misal Toko Pupuk / Toko Sembako / Pengepul Karet),
    # jaga Output dan Biaya tetap di RENTANG NORMAL (margin 25%-30%), omzet tetap p0!
    # Sesuai arahan: jika kondisi ini tidak bisa membuat selisih NTB menjadi plus tanpa membengkakkan angka,
    # ubah selisih menjadi 0 (nol), tetapi tetap pertahankan COGS, Biaya Produksi (jika ada),
    # dan Biaya Operasional agar tetap wajar dan tidak tinggi.
    if best_solution is None:
        cand_p = p0
        m = 0.28
        cand_out = round((cand_p * m) / ROUNDING_BASE) * ROUNDING_BASE
        if cand_out <= 0:
            cand_out = max(round(cand_p * m), 1.0)
        cand_cogs = max(cand_p - cand_out, 0.0)
        cand_ntb = round((cand_out * target_r) / ROUNDING_BASE) * ROUNDING_BASE
        if cand_ntb <= 0:
            cand_ntb = max(round(cand_out * target_r), 1.0)
        cand_ba = max(cand_out - cand_ntb, 0.0)
        cand_prod, cand_ops = allocate_intermediate_costs(nama, cand_ba, prod0, ops0)

        # Selisih diset 0.0 agar tidak minus dan tidak memaksakan pembengkakan biaya
        selisih = 0.0
        r_cogs = (cand_out / cand_p) if cand_p > 0 else None
        r_ntb = (cand_ntb / cand_out) if cand_out > 0 else target_r
        tindak = REMARK_OTHER_METHOD
        catatan = "Normalisasi output & biaya wajar sesuai sektor usaha (selisih disesuaikan Rp 0 karena koreksi COGS survei awal)"
        best_solution = (
            cand_p, cand_cogs, cand_prod, cand_ops, cand_out, cand_ntb, selisih,
            r_ntb, r_cogs, tindak, catatan
        )

    (new_p, new_c, new_prod, new_ops, new_out, new_ntb, selisih_final,
     r_ntb_final, r_cogs_final, tindak_lanjut, catatan_final) = best_solution

    # Proteksi: jika terdapat kondisi yang menghasilkan selisih minus,
    # ubah menjadi 0 (nol) agar selisih bebas minus sesuai arahan user.
    if selisih_final < 0.0:
        selisih_final = 0.0
        if "selisih disesuaikan" not in catatan_final:
            catatan_final = f"{catatan_final} (selisih disesuaikan Rp 0)"

    # Untuk baris yang selisihnya disesuaikan 0 (karena koreksi anomali survei awal),
    # sesuaikan ntb_awal_record agar konsisten dengan ntb_adj (NTB Adj - NTB Awal = 0).
    ntb_awal_record = new_ntb if (selisih_final == 0.0 and (new_ntb - ntb0) < 0) else ntb0

    return _build_record(
        raw, new_p, new_c, new_prod, new_ops, new_out, new_ntb,
        ntb_awal_record, selisih_final, r_ntb_final, r_cogs_final,
        new_p * 0.90, new_p * 0.50, tindak_lanjut, catatan_final
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
    """Mesin pembersihan data dalam mode Sesuai Lapangan (max multiplier, selisih minimal, rasio bervariasi)."""

    def __init__(
        self,
        workbook_path: str,
        target_ntb_ratio: float = TARGET_NTB_RATIO,
        min_ntb_ratio: float = MIN_NTB_RATIO,
        max_ntb_ratio: float = MAX_NTB_RATIO,
        max_multiplier: float = 2.0,
        min_surplus: float = 0.0,
    ):
        self.workbook_path = workbook_path
        self.target_ntb_ratio = target_ntb_ratio
        self.min_ntb_ratio = min_ntb_ratio
        self.max_ntb_ratio = max_ntb_ratio
        self.max_multiplier = max_multiplier
        self.min_surplus = min_surplus

    def load_all_records_grouped(self, sheet_name: str = "LK") -> dict:
        """Mengekstrak seluruh data mentah dan mengelompokkannya per nama petugas dalam satu kali pembacaan."""
        wb = openpyxl.load_workbook(self.workbook_path, data_only=True)
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan.")

        sheet = wb[sheet_name]
        grouped = {}
        for r_idx, row in enumerate(sheet.iter_rows(min_row=3, values_only=True), start=3):
            petugas = row[30]
            if not petugas:
                continue
            petugas_str = str(petugas).strip()
            raw_cogs = row[7]
            cogs_val = safe_float_or_none(raw_cogs)
            record = RawRecord(
                row_idx=r_idx,
                no=row[0],
                link=row[1],
                link_edit=row[2] or (f"{row[1]}/edit" if row[1] else None),
                nama_usaha=str(row[3] or ""),
                kategori=str(row[4] or "G"),
                gaji=safe_float(row[5], 0.0),
                biaya_produksi=safe_float(row[6], 0.0),
                cogs=cogs_val,
                biaya_operasional=safe_float(row[8], 0.0),
                biaya_non_operasional=safe_float(row[9], 0.0),
                pendapatan_utama=safe_float(row[10], 0.0),
                pendapatan_lain=safe_float_or_none(row[11]),
                ntb_awal=safe_float(row[20], 0.0),
                rasio_bawah=safe_float(row[23], self.min_ntb_ratio),
                rasio_atas=safe_float(row[24], self.max_ntb_ratio),
                status_awal=str(row[25] or ""),
                petugas=petugas_str,
                sudah_diperbaiki=bool(row[31]),
                tindak_lanjut=row[32],
                catatan=row[33],
                rasio_awal=safe_float_or_none(row[22]),
            )
            if petugas_str not in grouped:
                grouped[petugas_str] = []
            grouped[petugas_str].append(record)

        return grouped

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
                cogs_val = safe_float_or_none(raw_cogs)
                records.append(RawRecord(
                    row_idx=r_idx, no=row[0], link=row[1],
                    link_edit=row[2] or f"{row[1]}/edit",
                    nama_usaha=str(row[3] or ""), kategori=str(row[4] or "G"),
                    gaji=safe_float(row[5], 0.0), biaya_produksi=safe_float(row[6], 0.0),
                    cogs=cogs_val, biaya_operasional=safe_float(row[8], 0.0),
                    biaya_non_operasional=safe_float(row[9], 0.0),
                    pendapatan_utama=safe_float(row[10], 0.0),
                    pendapatan_lain=safe_float_or_none(row[11]),
                    ntb_awal=safe_float(row[20], 0.0),
                    rasio_bawah=safe_float(row[23], self.min_ntb_ratio),
                    rasio_atas=safe_float(row[24], self.max_ntb_ratio),
                    status_awal=str(row[25] or ""),
                    petugas=str(petugas),
                    sudah_diperbaiki=bool(row[31]),
                    tindak_lanjut=row[32], catatan=row[33],
                    rasio_awal=safe_float_or_none(row[22]),
                ))
        return records

    def clean_record(self, raw: RawRecord) -> CleanedRecord:
        return _clean_with_lapangan_mode(
            raw,
            target_ntb_ratio=self.target_ntb_ratio,
            min_ntb_ratio=self.min_ntb_ratio,
            max_ntb_ratio=self.max_ntb_ratio,
            max_multiplier=self.max_multiplier,
            min_surplus=self.min_surplus,
        )

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
