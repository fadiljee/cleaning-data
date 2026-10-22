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

    def test_contextual_cost_allocation_trade_vs_prod(self):
        """Uji alokasi biaya: komoditas tertutup murni (pangkalan gas) prod = 0, toko sembako/bengkel prod tetap ada"""
        raw_gas = self.create_sample_raw("DI ATAS 0.78", 120000000.0, 80000000.0, 5000000.0, 2000000.0, 25000000.0)
        raw_gas.nama_usaha = "PANGKALAN GAS ELPIJI (KARTINI)"
        cleaned_gas = self.cleaner_lapangan.clean_record(raw_gas)

        # Untuk pangkalan gas: biaya produksi fisik harus 0 (komoditas tertutup murni)
        self.assertEqual(cleaned_gas.biaya_produksi, 0.0)
        self.assertLess(cleaned_gas.biaya_operasional, 20000000.0)

        raw_trade = self.create_sample_raw("DI ATAS 0.78", 120000000.0, 80000000.0, 5000000.0, 2000000.0, 25000000.0)
        raw_trade.nama_usaha = "TOKO SEMBAKO MAJU JAYA"
        cleaned_trade = self.cleaner_lapangan.clean_record(raw_trade)
        # Untuk toko sembako yang di survei punya biaya kemasan/produksi (5M), pos produksi dipertahankan!
        self.assertGreater(cleaned_trade.biaya_produksi, 0.0)
        self.assertGreater(cleaned_trade.biaya_operasional, 0.0)

        raw_prod = self.create_sample_raw("DI ATAS 0.78", 120000000.0, 60000000.0, 8000000.0, 2000000.0, 30000000.0)
        raw_prod.nama_usaha = "BENGKEL MOTOR SENTOSA"
        cleaned_prod = self.cleaner_lapangan.clean_record(raw_prod)
        self.assertGreater(cleaned_prod.biaya_produksi, 0.0)
        self.assertGreater(cleaned_prod.biaya_operasional, 0.0)

    def test_varied_ntb_ratios_not_uniform(self):
        """Uji bahwa rasio NTB bervariasi alami antar usaha (tidak seragam 72.00%)"""
        ratios = []
        for idx in range(10):
            raw = self.create_sample_raw("DI ATAS 0.78", 100000000.0, 70000000.0, 0.0, 1000000.0, 20000000.0)
            raw.row_idx = 100 + idx
            raw.nama_usaha = f"TOKO USAHA KE-{idx}"
            cleaned = self.cleaner_lapangan.clean_record(raw)
            ratios.append(round(cleaned.rasio_ntb_output, 4))

        # Rasio tidak boleh semua persis sama
        unique_ratios = set(ratios)
        self.assertGreater(len(unique_ratios), 1)

    def test_normal_output_and_reasonable_selisih_for_topup(self):
        """Uji normalisasi output dan selisih wajar untuk usaha top up ewallet / cell"""
        raw = RawRecord(
            row_idx=3510,
            no=10.0,
            link="https://fasih.bps.go.id/test3510",
            link_edit="https://fasih.bps.go.id/test3510/edit",
            nama_usaha="TOP UP E-WALLET",
            kategori="G",
            gaji=0.0,
            biaya_produksi=28000000.0,
            cogs=None,
            biaya_operasional=10000000.0,
            biaya_non_operasional=0.0,
            pendapatan_utama=120000000.0,
            pendapatan_lain=None,
            ntb_awal=11500000.0,
            rasio_bawah=0.64,
            rasio_atas=0.78,
            status_awal="DALAM RENTANG ACUAN",
            petugas="Fadil Julianto",
            sudah_diperbaiki=False,
            tindak_lanjut=None,
            catatan=None,
            rasio_awal=0.6833
        )
        cleaned = self.cleaner_lapangan.clean_record(raw)

        # Output harus di rentang normal (bukan 120 juta)
        self.assertLess(cleaned.output, 60000000.0)
        self.assertGreaterEqual(cleaned.output, 12000000.0)

        # Selisih NTB tidak boleh 70 juta (harus wajar/minimal)
        self.assertLess(cleaned.selisih_ntb, 15000000.0)
        self.assertGreaterEqual(cleaned.selisih_ntb, 0.0)

        # Biaya produksi harus 0 untuk usaha loket/top up
        self.assertEqual(cleaned.biaya_produksi, 0.0)
        self.assertLess(cleaned.biaya_operasional, 15000000.0)

    def test_raw_anomaly_missing_cogs_selisih_zero_reasonable_costs(self):
        """Uji baris anomali data survei awal di mana COGS tertinggal/0 dan NTB0 sangat tinggi (e.g. Suhendri)"""
        # Omzet 120M, COGS awal 0, NTB0 113.76M (anomali survei awal)
        raw = RawRecord(
            row_idx=3518,
            no=18.0,
            link="https://fasih.bps.go.id/test3518",
            link_edit="https://fasih.bps.go.id/test3518/edit",
            nama_usaha="JUAL BELI KARET (SUHENDRI)",
            kategori="G",
            gaji=0.0,
            biaya_produksi=0.0,
            cogs=0.0,
            biaya_operasional=6240000.0,
            biaya_non_operasional=0.0,
            pendapatan_utama=120000000.0,
            pendapatan_lain=None,
            ntb_awal=113760000.0,
            rasio_bawah=0.64,
            rasio_atas=0.78,
            status_awal="DI ATAS 0.78",
            petugas="Fadil Julianto",
            sudah_diperbaiki=False,
            tindak_lanjut=None,
            catatan=None,
            rasio_awal=0.948
        )
        cleaned = self.cleaner_lapangan.clean_record(raw)

        # 1. Selisih NTB harus 0.0 (bebas minus)
        self.assertEqual(cleaned.selisih_ntb, 0.0)

        # 2. Output harus di rentang wajar sektor perdagangan (margin 20% - 35%)
        self.assertGreaterEqual(cleaned.output, 24000000.0)
        self.assertLessEqual(cleaned.output, 45000000.0)

        # 3. Biaya produksi harus 0 untuk pedagang karet
        self.assertEqual(cleaned.biaya_produksi, 0.0)

        # 4. Biaya operasional wajar dan tidak tinggi (< 15 juta)
        self.assertLessEqual(cleaned.biaya_operasional, 15000000.0)

        # 5. COGS wajar dan normal (> 70 juta)
        self.assertGreaterEqual(cleaned.cogs, 70000000.0)

        # 6. Rasio NTB tetap dalam rentang acuan hijau
        self.assertGreaterEqual(cleaned.rasio_ntb_output, MIN_NTB_RATIO)
        self.assertLessEqual(cleaned.rasio_ntb_output, MAX_NTB_RATIO)


if __name__ == "__main__":
    unittest.main()
