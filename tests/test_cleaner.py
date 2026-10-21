"""
Unit Tests untuk Data Cleaning & Validasi NTB
Memverifikasi kepatuhan seluruh batasan bisnis.
"""
import unittest
from src.cleaner import DataCleaner
from src.cleaner_lapangan import DataCleanerLapangan
from src.models import RawRecord
from src.config import MIN_NTB_RATIO, MAX_NTB_RATIO, MIN_COGS_MARGIN, MAX_COGS_MARGIN


class TestDataCleaner(unittest.TestCase):
    """Pengujian logika matematis dan batasan kualitas data"""

    def setUp(self):
        self.cleaner = DataCleaner("dummy.xlsx")
        self.cleaner_lapangan = DataCleanerLapangan("dummy.xlsx")

    def create_sample_raw(self, status: str, pend: float, cogs: float, prod: float, ops: float, ntb_awal: float) -> RawRecord:
        return RawRecord(
            row_idx=100,
            no=1.0,
            link="https://fasih.bps.go.id/test",
            link_edit="https://fasih.bps.go.id/test/edit",
            nama_usaha="TOKO CONTOH",
            kategori="G",
            gaji=0.0,
            biaya_produksi=prod,
            cogs=cogs,
            biaya_operasional=ops,
            biaya_non_operasional=0.0,
            pendapatan_utama=pend,
            pendapatan_lain=None,
            ntb_awal=ntb_awal,
            rasio_bawah=0.64,
            rasio_atas=0.78,
            status_awal=status,
            petugas="Fadil Julianto",
            sudah_diperbaiki=False,
            tindak_lanjut=None,
            catatan=None,
            rasio_awal=0.90
        )

    def test_clean_high_ratio_case(self):
        """Uji kasus rasio awal terlalu tinggi (DI ATAS 0.78) pada usaha perdagangan ritel"""
        # Misal omzet 120M, COGS 75M, Prod 0, Ops 2M, NTB awal 43M (rasio NTB awal ~0.95 di atas 0.78)
        raw = self.create_sample_raw("DI ATAS 0.78", 120000000.0, 75000000.0, 0.0, 2000000.0, 43000000.0)
        cleaned = self.cleaner.clean_record(raw)

        # 1. Rasio NTB harus masuk rentang acuan 0.64 - 0.78
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)

        # 2. Rasio COGS harus masuk rentang hijau 10% - 50%
        self.assertIsNotNone(cleaned.rasio_cogs)
        self.assertGreaterEqual(cleaned.rasio_cogs, MIN_COGS_MARGIN)
        self.assertLessEqual(cleaned.rasio_cogs, MAX_COGS_MARGIN)

        # 3. Selisih NTB tidak boleh minus dan untung sedikit
        self.assertGreater(cleaned.selisih_ntb, 0.0)

        # 4. Sudah diperbaiki = True
        self.assertTrue(cleaned.sudah_diperbaiki)
        self.assertEqual(cleaned.status_rasio, "DALAM RENTANG ACUAN")

    def test_clean_low_ratio_case(self):
        """Uji kasus rasio awal terlalu rendah (DI BAWAH 0.64) karena salah pos belanja"""
        # Misal omzet 120M, COGS 60M, Prod 35M (salah pos belanja stok), Ops 1.2M, NTB awal 23.8M
        raw = self.create_sample_raw("DI BAWAH 0.64", 120000000.0, 60000000.0, 35000000.0, 1200000.0, 23800000.0)
        cleaned = self.cleaner.clean_record(raw)

        # 1. Rasio NTB harus masuk rentang acuan 0.64 - 0.78
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)

        # 2. Rasio COGS harus masuk rentang hijau 10% - 50%
        self.assertIsNotNone(cleaned.rasio_cogs)
        self.assertGreaterEqual(cleaned.rasio_cogs, MIN_COGS_MARGIN)
        self.assertLessEqual(cleaned.rasio_cogs, MAX_COGS_MARGIN)

        # 3. Selisih NTB tidak boleh minus dan untung sedikit
        self.assertGreater(cleaned.selisih_ntb, 0.0)

        # 4. Status perbaikan
        self.assertTrue(cleaned.sudah_diperbaiki)
        self.assertEqual(cleaned.status_rasio, "DALAM RENTANG ACUAN")


    def test_clean_no_cogs_service_case(self):
        """Uji kasus usaha tanpa COGS karena salah penarikan KBLI (e.g. TOP UP E-WALLET)"""
        raw = RawRecord(
            row_idx=101,
            no=2.0,
            link="https://fasih.bps.go.id/test2",
            link_edit="https://fasih.bps.go.id/test2/edit",
            nama_usaha="TOP UP E-WALLET & PULSA",
            kategori="G",
            gaji=0.0,
            biaya_produksi=0.0,
            cogs=None,
            biaya_operasional=1000000.0,
            biaya_non_operasional=0.0,
            pendapatan_utama=15000000.0,
            pendapatan_lain=None,
            ntb_awal=10000000.0,
            rasio_bawah=0.64,
            rasio_atas=0.78,
            status_awal="DALAM RENTANG ACUAN",
            petugas="Fadil Julianto",
            sudah_diperbaiki=False,
            tindak_lanjut=None,
            catatan=None,
            rasio_awal=0.67
        )
        cleaned = self.cleaner.clean_record(raw)

        # 1. COGS dan Rasio COGS harus None (tidak dipaksa formula dagang)
        self.assertIsNone(cleaned.cogs)
        self.assertIsNone(cleaned.rasio_cogs)

        # 2. Pendapatan tidak membengkak (maksimal 1.5x)
        self.assertLessEqual(cleaned.pendapatan_utama, raw.pendapatan_utama * 1.5)

        # 3. Rasio NTB tetap dalam rentang acuan 0.64 - 0.78
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)

        # 4. Selisih NTB tidak boleh minus dan untung sedikit
        self.assertGreaterEqual(cleaned.selisih_ntb, 0.0)

    def test_multiplier_limit_and_positive_surplus(self):
        """Uji batasan multiplier <= 1.5x dan selisih NTB untung sedikit (tidak nol semua)"""
        raw = self.create_sample_raw("DI ATAS 0.78", 50000000.0, 20000000.0, 0.0, 1000000.0, 28000000.0)
        cleaned = self.cleaner.clean_record(raw)

        # 1. Multiplier tidak boleh melebihi 1.5x dari omzet awal
        multiplier = cleaned.pendapatan_utama / raw.pendapatan_utama
        self.assertLessEqual(multiplier, 1.50)

        # 2. Selisih NTB harus positif (ada untung sedikit, bukan 0)
        self.assertGreater(cleaned.selisih_ntb, 0.0)

    def test_cleaner_lapangan_mode(self):
        """Uji mode lapangan: multiplier <= 2x, selisih >= 0 (boleh 0 atau positif), rasio hijau"""
        raw = self.create_sample_raw("DI ATAS 0.78", 120000000.0, 75000000.0, 0.0, 2000000.0, 43000000.0)
        cleaned = self.cleaner_lapangan.clean_record(raw)

        multiplier = cleaned.pendapatan_utama / raw.pendapatan_utama
        self.assertLessEqual(multiplier, 2.00)
        self.assertGreaterEqual(cleaned.selisih_ntb, 0.0)
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)
        if cleaned.rasio_cogs is not None:
            self.assertGreaterEqual(cleaned.rasio_cogs, MIN_COGS_MARGIN)
            self.assertLessEqual(cleaned.rasio_cogs, MAX_COGS_MARGIN)


if __name__ == "__main__":
    unittest.main()
