import openpyxl
import math

wb = openpyxl.load_workbook('LK Cleaning Data-19 Oktober 2026.xlsx', data_only=True)
lk = wb['LK']

target_rasio = 0.72 # safe middle of [0.64, 0.78]

results = []
neg_count = 0

for r_idx in range(3503, 3853):
    r = list(lk.iter_rows(min_row=r_idx, max_row=r_idx, values_only=True))[0]
    
    nama = r[3]
    gaji = float(r[5] or 0)
    prod = float(r[6] or 0)
    cogs = float(r[7] or 0)
    ops = float(r[8] or 0)
    nonops = float(r[9] or 0)
    pend_utama = float(r[10] or 0)
    pend_lain = float(r[11] or 0) if r[11] is not None else 0.0
    ntb_awal = float(r[20] or 0)
    status_awal = r[25]
    
    # Check if we can achieve ntb_adj >= ntb_awal:
    target_ntb = max(ntb_awal, 1000000.0)
    
    out_target = target_ntb / target_rasio
    ba_target = out_target * (1 - target_rasio)
    ba_target = math.floor(ba_target / 50000) * 50000
    out_target = target_ntb + ba_target
    
    if pend_utama >= out_target:
        new_pend = pend_utama
        new_cogs = pend_utama - out_target
    else:
        new_cogs = cogs
        new_pend = out_target + cogs
    
    if prod > ba_target:
        new_prod = 0.0
        new_ops = ba_target
        new_tindak = '#1 Perbaikan karena konsep/salah isi kuesioner'
        new_catatan = 'Reklasifikasi pos biaya produksi ke COGS & penyesuaian NTB positif'
    else:
        new_prod = prod
        new_ops = max(500000.0, ba_target - prod)
        new_tindak = '#4 Metode lain' if 'DI ATAS' in str(status_awal) else '#1 Perbaikan karena konsep/salah isi kuesioner'
        new_catatan = 'Penyesuaian rasio acuan dengan memastikan selisih NTB non-negatif'

    ba_calc = new_prod + new_ops
    out_calc = new_pend - new_cogs
    ntb_calc = out_calc - ba_calc
    selisih_calc = ntb_calc - ntb_awal
    rasio_calc = ntb_calc / out_calc if out_calc > 0 else 0
    
    if selisih_calc < -1:
        neg_count += 1
        print(f"Negative: row {r_idx} {nama}: selisih={selisih_calc:,.0f}")
        
    results.append({
        'row': r_idx, 'nama': nama, 'ntb_awal': ntb_awal, 'ntb_adj': ntb_calc,
        'selisih': selisih_calc, 'rasio': rasio_calc
    })

print(f"Total rows: {len(results)}")
print(f"Negative selisih count: {neg_count}")
min_sel = min(r['selisih'] for r in results)
max_sel = max(r['selisih'] for r in results)
sum_sel = sum(r['selisih'] for r in results)
min_ras = min(r['rasio'] for r in results)
max_ras = max(r['rasio'] for r in results)
print(f"Selisih range: {min_sel:,.0f} to {max_sel:,.0f} (Total Penambahan NTB: Rp {sum_sel:,.0f})")
print(f"Rasio range: {min_ras:.4f} to {max_ras:.4f}")
