"""
Konfigurasi dan Konstanta Bisnis untuk Data Cleaning LK
"""
from pathlib import Path

# Paths
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"

DEFAULT_INPUT_FILE = "LK Cleaning Data-19 Oktober 2026.xlsx"

# Rasio Acuan Ekonomi Kategori G (Perdagangan)
TARGET_NTB_RATIO = 0.72       # Titik tengah aman dari rentang acuan [0.64, 0.78]
MIN_NTB_RATIO = 0.64          # Batas bawah rasio NTB / Output
MAX_NTB_RATIO = 0.78          # Batas atas rasio NTB / Output

# Rasio Acuan Margin COGS (Status Berdasarkan Warna)
TARGET_COGS_MARGIN = 0.30     # Titik tengah aman margin kotor [0.10, 0.50]
MIN_COGS_MARGIN = 0.10        # Batas bawah margin kotor (10% - Hijau)
MAX_COGS_MARGIN = 0.50        # Batas atas margin kotor (50% - Hijau)

# Parameter Pembulatan Rupiah
ROUNDING_BASE = 10000         # Pembulatan biaya antara ke kelipatan Rp 10.000

# Kode Remark Tindak Lanjut
REMARK_NO_CHANGE = "#0 Tidak ada perbaikan"
REMARK_CONCEPT_FIX = "#1 Perbaikan karena konsep/salah isi kuesioner"
REMARK_OTHER_METHOD = "#4 Metode lain"
