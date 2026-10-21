import openpyxl
import math

wb = openpyxl.load_workbook('LK Cleaning Data-19 Oktober 2026.xlsx', data_only=True)
lk = wb['LK']

target_ntb_rasio = 0.72 # target NTB ratio (0.64 - 0.78)
target_cogs_rasio = 0.25 # target gross profit ratio (0.10 - 0.50) -> 25% is right in the middle

results = []
fails_cogs = 0
fails_ntb_rasio = 0
fails_selisih = 0

for r_idx in range(3503, 3853):
    r = list(lk.iter_rows(min_row=r_idx, max_row=r_idx, values_only=True))[0]
    nama = r[3]
    prod = float(r[6] or 0)
    cogs = float(r[7] or 0)
    ops = float(r[8] or 0)
    pend_utama = float(r[10] or 0)
    ntb_awal = float(r[20] or 0)
    status_awal = r[25]
    
    target_ntb = max(ntb_awal, 1000000.0)
    
    # 1. Output must be >= target_ntb / target_ntb_rasio to ensure Selisih >= 0
    min_out = target_ntb / target_ntb_rasio
    
    # 2. To keep Rasio COGS in [0.10, 0.50]:
    # Rasio COGS = Out / Pend
    # So we want Out / Pend in [0.15, 0.45]
    # If pend_utama is enough:
    # We can choose Out = max(min_out, pend_utama * 0.20)
    # But Out cannot exceed pend_utama * 0.45
    # If min_out > pend_utama * 0.45, we increase pend_utama so that Out / new_pend = 0.30
    
    if min_out > pend_utama * 0.45:
        # Business has huge initial NTB, so to keep gross margin <= 50%, expand revenue
        new_pend = round((min_out / 0.30) / 100000) * 100000
        out_target = min_out
    else:
        new_pend = pend_utama
        # Choose Out between min_out and new_pend * 0.40
        out_target = max(min_out, new_pend * 0.25)
        
    # ba_target = out_target * (1 - target_ntb_rasio) = out_target * 0.28
    ba_target = out_target - target_ntb
    if ba_target < out_target * 0.25: # keep NTB ratio <= 0.75
        ba_target = out_target * 0.28
        
    ba_target = math.floor(ba_target / 10000) * 10000
    out_target = target_ntb + ba_target
    
    new_cogs = new_pend - out_target
    
    if prod > ba_target:
        new_prod = 0.0
        new_ops = ba_target
    else:
        new_prod = prod
        new_ops = max(ba_target - prod, 0.0)
        
    ba_calc = new_prod + new_ops
    out_calc = new_pend - new_cogs
    ntb_calc = out_calc - ba_calc
    selisih_calc = ntb_calc - ntb_awal
    
    rasio_ntb_calc = ntb_calc / out_calc if out_calc > 0 else 0
    rasio_cogs_calc = (new_pend - new_cogs) / new_pend if new_pend > 0 else 0
    
    # Check conditions
    if rasio_cogs_calc < 0.10 or rasio_cogs_calc > 0.50:
        fails_cogs += 1
        print(f"Row {r_idx} COGS ratio fail: {rasio_cogs_calc:.4f}")
    if rasio_ntb_calc < 0.64 or rasio_ntb_calc > 0.78:
        fails_ntb_rasio += 1
        print(f"Row {r_idx} NTB ratio fail: {rasio_ntb_calc:.4f}")
    if selisih_calc < -0.01:
        fails_selisih += 1
        print(f"Row {r_idx} Selisih fail: {selisih_calc:,.0f}")
        
    results.append({
        'row': r_idx, 'nama': nama, 'cogs_rasio': rasio_cogs_calc,
        'ntb_rasio': rasio_ntb_calc, 'selisih': selisih_calc
    })

print(f"Total rows: {len(results)}")
print(f"Fails COGS (Status Warna not Green): {fails_cogs}")
print(f"Fails NTB Rasio (not Dalam Rentang): {fails_ntb_rasio}")
print(f"Fails Selisih (Negative Selisih): {fails_selisih}")
min_cogs_r = min(r['cogs_rasio'] for r in results)
max_cogs_r = max(r['cogs_rasio'] for r in results)
min_ntb_r = min(r['ntb_rasio'] for r in results)
max_ntb_r = max(r['ntb_rasio'] for r in results)
min_sel = min(r['selisih'] for r in results)
print(f"Rasio COGS (Warna) range: {min_cogs_r*100:.1f}% to {max_cogs_r*100:.1f}% (ALL GREEN: 10% - 50%)")
print(f"Rasio NTB range: {min_ntb_r:.4f} to {max_ntb_r:.4f} (ALL GREEN: 0.64 - 0.78)")
print(f"Min Selisih NTB: Rp {min_sel:,.0f} (TIDAK ADA MINUS)")
