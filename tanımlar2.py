import pandas as pd
import numpy as np
import time
from tqdm import tqdm
from ortools.linear_solver import pywraplp

# Parametreler
SIMULASYON_SAYISI = 12
np.random.seed(12)

# Excel'den veri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))

satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# === Talep Senaryolarını üret (progress bar ile) ===
print("Talep senaryoları oluşturuluyor...")
sales_scenarios = {}

for urun in tqdm(urunler, desc="Senaryo Üretimi", unit="ürün"):
    ort = urun_param_dict[urun]["ortalama"]
    std = urun_param_dict[urun]["std"]
    sales_scenarios[urun] = [max(0, int(round(np.random.normal(ort, std)))) for _ in range(SIMULASYON_SAYISI)]

# === Modelleme ===
solver = pywraplp.Solver.CreateSolver('SCIP')

uretim_miktarlari = {
    urun: solver.IntVar(0, max(sales_scenarios[urun]) * 2, f'uretim_{urun}') for urun in urunler
}

x = {
    (urun, uretici): solver.IntVar(0, max(sales_scenarios[urun]) * 2, f'x_{urun}_{uretici}')
    for urun in urunler for uretici in ureticiler if (urun, uretici) in urun_uretici_dict
}

satis_vars = {
    (urun, s): solver.IntVar(0, max(sales_scenarios[urun]), f'satis_{urun}_{s}')
    for urun in urunler for s in range(SIMULASYON_SAYISI)
}

# Kısıtlar
for urun in urunler:
    solver.Add(
        sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x) == uretim_miktarlari[urun]
    )

for uretici in ureticiler:
    toplam = sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x)
    solver.Add(toplam <= uretici_kapasite_dict.get(ureticiler, float('inf')))
    solver.Add(toplam >= uretici_alt_kapasite_dict.get(ureticiler, 0))

for urun in urunler:
    for s in range(SIMULASYON_SAYISI):
        talep = sales_scenarios[urun][s]
        satis = satis_vars[(urun, s)]
        solver.Add(satis <= talep)
        solver.Add(satis <= uretim_miktarlari[urun])

# Amaç fonksiyonu
objective = solver.Objective()

for urun in urunler:
    for s in range(SIMULASYON_SAYISI):
        objective.SetCoefficient(satis_vars[(urun, s)], satis_fiyat[urun] / SIMULASYON_SAYISI)

for (urun, uretici), var in x.items():
    maliyet = urun_uretici_dict[(urun, uretici)]
    objective.SetCoefficient(var, -maliyet)

objective.SetMaximization()

# === Modeli çöz (süre ölçümü dahil) ===
print("Model çözülüyor...")
start_time = time.time()
status = solver.Solve()
end_time = time.time()
elapsed = end_time - start_time

if status == pywraplp.Solver.OPTIMAL:
    print("\n=== Optimum Üretim Miktarları ===")
    toplam_gelir = 0
    toplam_gider = 0

    for urun in urunler:
        print(f"{urun}: {uretim_miktarlari[urun].solution_value():.0f} adet üretilecek")
        gelir = sum(satis_vars[(urun, s)].solution_value() * satis_fiyat[urun] for s in range(SIMULASYON_SAYISI)) / SIMULASYON_SAYISI
        toplam_gelir += gelir

    for (urun, uretici), var in x.items():
        toplam_gider += var.solution_value() * urun_uretici_dict[(urun, uretici)]

    net_kar = toplam_gelir - toplam_gider
    print(f"\nToplam Gelir (Beklenen): {toplam_gelir:,.2f}")
    print(f"Toplam Gider: {toplam_gider:,.2f}")
    print(f"Net Beklenen Kar: {net_kar:,.2f}")
    print(f"Çözüm süresi: {elapsed:.2f} saniye")
else:
    print("Model çözülemedi.")
