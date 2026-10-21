"""
Model Data Struktur untuk Data Cleaning LK
"""
from dataclasses import dataclass
from typing import Optional


@dataclass
class RawRecord:
    """Representasi data baris awal dari sheet LK"""
    row_idx: int
    no: Optional[float]
    link: Optional[str]
    link_edit: Optional[str]
    nama_usaha: str
    kategori: str
    gaji: float
    biaya_produksi: float
    cogs: Optional[float]
    biaya_operasional: float
    biaya_non_operasional: float
    pendapatan_utama: float
    pendapatan_lain: Optional[float]
    ntb_awal: float
    rasio_bawah: float
    rasio_atas: float
    status_awal: str
    petugas: str
    sudah_diperbaiki: bool
    tindak_lanjut: Optional[str]
    catatan: Optional[str]
    rasio_awal: Optional[float] = None


@dataclass
class CleanedRecord:
    """Representasi data baris yang telah dibersihkan dan disesuaikan"""
    row_idx: int
    no: Optional[float]
    link: Optional[str]
    link_edit: Optional[str]
    nama_usaha: str
    kategori: str
    gaji: float
    biaya_produksi: float
    cogs: Optional[float]
    biaya_operasional: float
    biaya_non_operasional: float
    pendapatan_utama: float
    pendapatan_lain: Optional[float]
    pendapatan_per_tahun: float
    pendapatan_per_bulan: float
    rasio_cogs: Optional[float]
    cogs_bawah: Optional[float]
    cogs_atas: Optional[float]
    biaya_antara: float
    output: float
    ntb_adj: float
    ntb_awal: float
    selisih_ntb: float
    rasio_ntb_output: float
    bawah: float
    atas: float
    status_rasio: str
    ba_bawah: float
    ba_atas: float
    pend_bawah: float
    pend_atas: float
    nama_petugas: str
    sudah_diperbaiki: bool
    tindak_lanjut: str
    catatan: str
    status_awal: str
    rasio_awal: Optional[float]
    prod_awal: float
    cogs_awal: Optional[float]
    ops_awal: float
    pend_awal: float


@dataclass
class CleaningSummary:
    """Ringkasan metrik statistik hasil proses cleaning"""
    petugas: str
    total_usaha: int
    total_selesai: int
    total_warna_hijau: int
    total_rasio_hijau: int
    total_selisih_minus: int
    total_ntb_awal: float
    total_ntb_adj: float
    total_selisih_ntb: float
    min_cogs_ratio: float
    max_cogs_ratio: float
    min_ntb_ratio: float
    max_ntb_ratio: float
