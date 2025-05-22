import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
from tqdm import tqdm
import time

# === Parametreler ===
SIMULASYON_SAYISI = 10000  # Senaryo sayısı (Sample Average Approximation)
SCENARIO_PROB = 1 / SIMULASYON_SAYISI

# === Excel'den veri okuma ===
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

# === Veri yapıları ===
urunler = [u for u in urun_kisit_data['Ürün'] if u != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# === Senaryoları oluştur ===
np.random.seed(42)
sales_scenarios = {}
print("=== Senaryo Oluşturuluyor ===")
for urun in urunler:
    ortalama = urun_param_dict[urun]['ortalama']
    std = urun_param_dict[urun]['std']
    mu = np.log(ortalama**2 / np.sqrt(ortalama**2 + std**2))
    sigma = np.sqrt(np.log(1 + (std**2 / ortalama**2)))
    sales_scenarios[urun] = np.random.lognormal(mu, sigma, SIMULASYON_SAYISI)

# === Model başlatma ===
solver = pywraplp.Solver.CreateSolver('SCIP')
start_time = time.time()

# === Karar değişkenleri ===
x = {(u, j): solver.IntVar(0, solver.infinity(), f"x_{u}_{j}") for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict}
satilan = {(u, k): solver.NumVar(0, solver.infinity(), f"satilan_{u}_{k}") for u in urunler for k in range(SIMULASYON_SAYISI)}

# === Amaç fonksiyonu ===
objective = solver.Objective()
for (u, j), var in x.items():
    objective.SetCoefficient(var, -urun_uretici_dict[(u, j)])
for u in urunler:
    for k in range(SIMULASYON_SAYISI):
        objective.SetCoefficient(satilan[(u, k)], satis_fiyat[u] * SCENARIO_PROB)
objective.SetMaximization()

# === Kısıtlar ===
with tqdm(total=len(urunler), desc="Satış kısıtları") as pbar:
    for u in urunler:
        toplam_uretim = sum(x[(u, j)] for j in ureticiler if (u, j) in x)
        for k in range(SIMULASYON_SAYISI):
            solver.Add(satilan[(u, k)] <= toplam_uretim)
            solver.Add(satilan[(u, k)] <= sales_scenarios[u][k])
        pbar.update(1)

with tqdm(total=len(ureticiler), desc="Kapasite kısıtları") as pbar:
    for j in ureticiler:
        toplam_uretim = sum(x[(u, j)] for u in urunler if (u, j) in x)
        solver.Add(toplam_uretim <= uretici_kapasite_dict.get(j, float('inf')))
        solver.Add(toplam_uretim >= uretici_alt_kapasite_dict.get(j, 0))
        pbar.update(1)

# === Model çözümü ===
print("\n=== Model Çözülüyor ===")
with tqdm(total=100, desc="Optimizasyon") as pbar:
    pbar.update(10)
    status = solver.Solve()
    pbar.update(90)

total_runtime = time.time() - start_time
print(f"Toplam çalışma süresi: {total_runtime:.2f} saniye")

# === Sonuçlar ===
if status == pywraplp.Solver.OPTIMAL:
    toplam_maliyet = sum(x[(u, j)].solution_value() * urun_uretici_dict[(u, j)] for (u, j) in x)
    beklenen_gelir = sum(satilan[(u, k)].solution_value() * satis_fiyat[u] * SCENARIO_PROB for u in urunler for k in range(SIMULASYON_SAYISI))
    beklenen_kar = beklenen_gelir - toplam_maliyet
    print("\n=== ÇÖZÜM SONUÇLARI ===")
    print(f"Toplam Üretim Maliyeti: {toplam_maliyet:,.2f}")
    print(f"Beklenen Satış Geliri: {beklenen_gelir:,.2f}")
    print(f"Beklenen Kar: {beklenen_kar:,.2f}")
else:
    print("Model optimal çözüme ulaşamadı. Durum kodu:", status)
