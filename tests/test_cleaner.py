"""
Unit Tests untuk Data Cleaning & Validasi NTB
Memverifikasi kepatuhan seluruh batasan bisnis.
"""
import unittest
from src.cleaner import DataCleaner
from src.models import RawRecord
from src.config import MIN_NTB_RATIO, MAX_NTB_RATIO, MIN_COGS_MARGIN, MAX_COGS_MARGIN


class TestDataCleaner(unittest.TestCase):
    """Pengujian logika matematis dan batasan kualitas data"""

    def setUp(self):
        self.cleaner = DataCleaner("dummy.xlsx")

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
        """Uji kasus rasio awal terlalu tinggi (DI ATAS 0.78)"""
        # Misal omzet 120M, COGS 48M, Prod 0, Ops 2M, NTB awal 70M
        raw = self.create_sample_raw("DI ATAS 0.78", 120000000.0, 48000000.0, 0.0, 2000000.0, 70000000.0)
        cleaned = self.cleaner.clean_record(raw)

        # 1. Rasio NTB harus masuk rentang acuan 0.64 - 0.78
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)

        # 2. Rasio COGS harus masuk rentang hijau 10% - 50%
        self.assertGreaterEqual(cleaned.rasio_cogs, MIN_COGS_MARGIN)
        self.assertLessEqual(cleaned.rasio_cogs, MAX_COGS_MARGIN)

        # 3. Selisih NTB tidak boleh minus
        self.assertGreaterEqual(cleaned.selisih_ntb, -0.01)

        # 4. Sudah diperbaiki = True
        self.assertTrue(cleaned.sudah_diperbaiki)
        self.assertEqual(cleaned.status_rasio, "DALAM RENTANG ACUAN")

    def test_clean_low_ratio_case(self):
        """Uji kasus rasio awal terlalu rendah (DI BAWAH 0.64) karena salah pos belanja"""
        # Misal omzet 120M, COGS 22M, Prod 50M (salah pos), Ops 1.2M, NTB awal 46.8M
        raw = self.create_sample_raw("DI BAWAH 0.64", 120000000.0, 22000000.0, 50000000.0, 1200000.0, 46800000.0)
        cleaned = self.cleaner.clean_record(raw)

        # 1. Rasio NTB harus masuk rentang acuan 0.64 - 0.78
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)

        # 2. Rasio COGS harus masuk rentang hijau 10% - 50%
        self.assertGreaterEqual(cleaned.rasio_cogs, MIN_COGS_MARGIN)
        self.assertLessEqual(cleaned.rasio_cogs, MAX_COGS_MARGIN)

        # 3. Selisih NTB tidak boleh minus
        self.assertGreaterEqual(cleaned.selisih_ntb, -0.01)

        # 4. Status perbaikan
        self.assertTrue(cleaned.sudah_diperbaiki)
        self.assertEqual(cleaned.status_rasio, "DALAM RENTANG ACUAN")


if __name__ == "__main__":
    unittest.main()
