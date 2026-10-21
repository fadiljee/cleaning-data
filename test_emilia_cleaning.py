import openpyxl
import math

wb = openpyxl.load_workbook('LK Cleaning Data-19 Oktober 2026.xlsx', data_only=True)
lk = wb['LK']

target_ntb_rasio = 0.72
target_cogs_rasio = 0.30

results = []
fails_cogs = 0
fails_ntb_rasio = 0
fails_selisih = 0

for r_idx in range(1403, 1753):
    r = list(lk.iter_rows(min_row=r_idx, max_row=r_idx, values_only=True))[0]
    nama = r[3]
    prod = float(r[6] or 0)
    cogs = float(r[7] or 0)
    ops = float(r[8] or 0)
    pend_utama = float(r[10] or 0)
    ntb_awal = float(r[20] or 0)
    status_awal = r[25]
    
    target_ntb = max(ntb_awal, 1000000.0)
    
    # 1. Output target
    out_target = target_ntb / target_ntb_rasio
    ba_target = out_target - target_ntb
    ba_target = math.floor(ba_target / 10000) * 10000
    out_target = target_ntb + ba_target
    
    # 2. Check Rasio COGS
    cogs_r = out_target / pend_utama if pend_utama > 0 else 0
    if 0.12 <= cogs_r <= 0.48:
        new_pend = pend_utama
    elif cogs_r > 0.48:
        new_pend = round((out_target / target_cogs_rasio) / 100000) * 100000
    else:
        new_pend = round((out_target / 0.20) / 100000) * 100000
        
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
max_sel = max(r['selisih'] for r in results)
sum_sel = sum(r['selisih'] for r in results)
print(f"Rasio COGS (Warna) range: {min_cogs_r*100:.2f}% to {max_cogs_r*100:.2f}% (ALL GREEN: 10% - 50%)")
print(f"Rasio NTB range: {min_ntb_r:.4f} to {max_ntb_r:.4f} (ALL GREEN: 0.64 - 0.78)")
print(f"Selisih NTB range: Rp {min_sel:,.0f} to Rp {max_sel:,.0f} (TIDAK ADA MINUS)")
print(f"Total Surplus Penambahan NTB: Rp {sum_sel:,.0f}")
