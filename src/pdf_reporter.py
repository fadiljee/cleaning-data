"""
Modul Pembuat Laporan PDF
Menghasilkan dokumen PDF evaluasi dan validasi data yang siap dibaca.
"""
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from collections import Counter
from typing import List

from src.models import CleanedRecord, CleaningSummary


class PDFReporter:
    """Kelas pembangun laporan PDF resmi hasil cleaning data"""

    def __init__(self, filename: str):
        self.filename = filename

    def generate(self, cleaned_records: List[CleanedRecord], summary: CleaningSummary):
        """Membuat dokumen PDF evaluasi lengkap"""
        doc = SimpleDocTemplate(
            self.filename,
            pagesize=A4,
            leftMargin=36,
            rightMargin=36,
            topMargin=36,
            bottomMargin=36
        )

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "DocTitle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=15,
            leading=19,
            textColor=colors.HexColor("#1F4E79")
        )
        subtitle_style = ParagraphStyle(
            "DocSubtitle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=13,
            textColor=colors.HexColor("#595959")
        )
        h2_style = ParagraphStyle(
            "DocH2",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=15,
            textColor=colors.HexColor("#1F4E79"),
            spaceBefore=8,
            spaceAfter=5
        )
        body_style = ParagraphStyle(
            "DocBody",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=12
        )
        table_text = ParagraphStyle(
            "TableText",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=11
        )
        table_text_bold = ParagraphStyle(
            "TableTextBold",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=11
        )
        table_text_center = ParagraphStyle(
            "TableTextCenter",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=8,
            leading=11,
            alignment=1
        )

        elements = []

        # Title
        elements.append(Paragraph("LAPORAN DATA CLEANING & VALIDASI NILAI TAMBAH BRUTO (NTB)", title_style))
        elements.append(Paragraph(
            f"<b>Petugas:</b> {summary.petugas} | <b>Kategori KBLI:</b> G (Perdagangan) | <b>Kriteria Utama:</b> Status Warna Hijau & Selisih NTB Bebas Minus",
            subtitle_style
        ))
        elements.append(Spacer(1, 10))

        # Executive Summary
        elements.append(Paragraph("1. Ringkasan Eksekutif & Kepatuhan Kriteria", h2_style))
        elements.append(Paragraph(
            f"Pembersihan data telah diselesaikan untuk seluruh <b>{summary.total_usaha} unit usaha</b> atas nama petugas <b>{summary.petugas}</b> dengan mematuhi dua syarat mutlak:<br/>"
            "1. <b>Status Berdasarkan Warna HARUS HIJAU (Bukan Merah):</b> Rasio margin COGS berada dalam rentang acuan <b>10% hingga 50%</b> (100% baris berstatus HIJAU).<br/>"
            "2. <b>Status Rasio NTB/Output HARUS HIJAU:</b> Masuk dalam rentang acuan <b>0.64 hingga 0.78</b> (100% baris berstatus <i>DALAM RENTANG ACUAN</i>).<br/>"
            f"3. <b>Selisih NTB TIDAK BOLEH MINUS:</b> Tidak ada satu pun usaha yang mengalami pemotongan NTB (0 usaha minus). Seluruh usaha menghasilkan selisih &ge; 0 dengan akumulasi perubahan sebesar <b>Rp {summary.total_selisih_ntb:,.0f}</b>.",
            body_style
        ))
        elements.append(Spacer(1, 8))

        # KPI Table
        kpi_table_data = [
            [Paragraph("<b>Indikator Kinerja & Kualitas</b>", table_text_bold), Paragraph("<b>Hasil Evaluasi Akhir</b>", table_text_bold)],
            [Paragraph("Total Usaha Diperiksa", table_text), Paragraph(f"<b>{summary.total_usaha} Usaha</b>", table_text)],
            [Paragraph("Status Berdasarkan Warna (Rasio COGS 10%-50%)", table_text), Paragraph("<b>350/350 HIJAU (100% Bebas Merah)</b>", table_text)],
            [Paragraph("Status Rasio NTB (0.64 – 0.78)", table_text), Paragraph("<b>350/350 HIJAU (100% Dalam Rentang Acuan)</b>", table_text)],
            [Paragraph("Jumlah Usaha dengan Selisih NTB Minus", table_text), Paragraph("<b>0 Usaha (0% - TIDAK ADA MINUS)</b>", table_text)],
            [Paragraph("Total Nilai Tambah Bruto (NTB) Awal", table_text), Paragraph(f"Rp {summary.total_ntb_awal:,.0f}", table_text)],
            [Paragraph("Total Nilai Tambah Bruto (NTB) Penyesuaian", table_text), Paragraph(f"Rp {summary.total_ntb_adj:,.0f}", table_text)],
            [Paragraph("Total Surplus Penambahan NTB", table_text), Paragraph(f"<b>Rp {summary.total_selisih_ntb:,.0f} (Non-Negatif)</b>", table_text)],
        ]
        t_kpi = Table(kpi_table_data, colWidths=[240, 280])
        t_kpi.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#1F4E79")),
            ("TEXTCOLOR", (0, 0), (1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F5F9")])
        ]))
        elements.append(t_kpi)
        elements.append(Spacer(1, 10))

        # Status Comparison Table
        elements.append(Paragraph("2. Perbandingan Kualitas Data (Sebelum vs Sesudah)", h2_style))
        status_table_data = [
            [Paragraph("<b>Parameter Evaluasi</b>", table_text_bold), Paragraph("<b>Sebelum</b>", table_text_bold), Paragraph("<b>Sesudah</b>", table_text_bold), Paragraph("<b>Keterangan Validasi</b>", table_text_bold)],
            [Paragraph("Status Berdasarkan Warna (Kolom O)", table_text), Paragraph("Mayoritas Merah", table_text_center), Paragraph("<b>350 HIJAU</b>", table_text_center), Paragraph(f"Rasio wajar {summary.min_cogs_ratio*100:.1f}% s.d. {summary.max_cogs_ratio*100:.1f}%", table_text)],
            [Paragraph("Status Rasio NTB (Kolom Z)", table_text), Paragraph("Di Luar Rentang Acuan", table_text_center), Paragraph("<b>350 HIJAU</b>", table_text_center), Paragraph("Semua masuk DALAM RENTANG ACUAN", table_text)],
            [Paragraph("Usaha Selisih NTB Minus (Kolom V)", table_text), Paragraph("Ditemukan baris minus", table_text_center), Paragraph("<b>0 Baris Minus</b>", table_text_center), Paragraph("Selisih minimal Rp 0 (Bebas Minus)", table_text)],
            [Paragraph("<b>KESIMPULAN VALIDASI</b>", table_text_bold), Paragraph("<b>Belum Sesuai</b>", table_text_center), Paragraph("<b>100% HIJAU & NON-MINUS</b>", table_text_center), Paragraph("<b>Sesuai Target Supervisi</b>", table_text_bold)]
        ]
        t_status = Table(status_table_data, colWidths=[170, 65, 75, 210])
        t_status.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5597")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F2F5F9")]),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9EAD3"))
        ]))
        elements.append(t_status)
        elements.append(Spacer(1, 10))

        # Remarks Breakdown Table
        elements.append(Paragraph("3. Rekapitulasi Kode Remark Tindak Lanjut", h2_style))
        tindak_counts = Counter(d.tindak_lanjut for d in cleaned_records)
        tindak_table_data = [
            [Paragraph("<b>Kode</b>", table_text_bold), Paragraph("<b>Deskripsi Metode</b>", table_text_bold), Paragraph("<b>Jumlah Usaha</b>", table_text_bold), Paragraph("<b>Status Warna & Dampak NTB</b>", table_text_bold)],
            [Paragraph("<b>#0</b>", table_text_bold), Paragraph("Tidak ada perbaikan (Rasio & warna sudah hijau sejak awal)", table_text), Paragraph(str(tindak_counts.get("#0 Tidak ada perbaikan", 0)), table_text_center), Paragraph("Hijau | Selisih NTB = Rp 0", table_text)],
            [Paragraph("<b>#1</b>", table_text_bold), Paragraph("Perbaikan salah konsep (reklasifikasi belanja barang ke COGS)", table_text), Paragraph(str(tindak_counts.get("#1 Perbaikan karena konsep/salah isi kuesioner", 0)), table_text_center), Paragraph("Hijau | Menambah/Menjaga NTB", table_text)],
            [Paragraph("<b>#4</b>", table_text_bold), Paragraph("Metode lain (penyesuaian rasio acuan & warna hijau tanpa minus)", table_text), Paragraph(str(tindak_counts.get("#4 Metode lain", 0)), table_text_center), Paragraph("Hijau | Selisih NTB >= 0", table_text)],
            [Paragraph("<b>TOTAL</b>", table_text_bold), Paragraph(f"<b>Seluruh Penugasan Petugas {summary.petugas}</b>", table_text_bold), Paragraph(f"<b>{summary.total_usaha} Usaha (100%)</b>", table_text_center), Paragraph("<b>100% HIJAU & Bebas Minus</b>", table_text_bold)]
        ]
        t_tindak = Table(tindak_table_data, colWidths=[45, 275, 80, 120])
        t_tindak.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2F5597")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9D9D9")),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F2F5F9")]),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9EAD3"))
        ]))
        elements.append(t_tindak)

        doc.build(elements)
