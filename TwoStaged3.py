import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp

# === Parametreler ===
SIMULASYON_SAYISI = 100  # Belirli sayıda senaryo
np.random.seed(12)

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

# === 1. AŞAMA: Senaryoları oluştur ===
sales_scenarios = {urun: [] for urun in urunler}
for k in range(SIMULASYON_SAYISI):
    np.random.seed(1000 + k)
    for urun in urunler:
        talep = max(0, np.random.normal(urun_param_dict[urun]['ortalama'], urun_param_dict[urun]['std']))
        sales_scenarios[urun].append(talep)

# === 2. AŞAMA: Senaryoları göz önünde bulundurarak optimal üretim kararlarını bul ===
solver = pywraplp.Solver.CreateSolver('SCIP')
x = {(u, j): solver.IntVar(0, solver.infinity(), f"x_{u}_{j}") for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict}

# Beklenen toplam karı maksimize et
objective = solver.Objective()
for (u, j), var in x.items():
    avg_talep = sum(sales_scenarios[u]) / SIMULASYON_SAYISI
    kar = (satis_fiyat[u] - urun_uretici_dict[(u, j)]) * avg_talep / SIMULASYON_SAYISI
    objective.SetCoefficient(var, kar)
objective.SetMaximization()

# Ürün ve üretici kapasite kısıtları
for u in urunler:
    solver.Add(sum(x[(u, j)] for j in ureticiler if (u, j) in x) <= max(sales_scenarios[u]))

for j in ureticiler:
    toplam = sum(x[(u, j)] for u in urunler if (u, j) in x)
    solver.Add(toplam <= uretici_kapasite_dict.get(j, float('inf')))
    solver.Add(toplam >= uretici_alt_kapasite_dict.get(j, 0))

solver.Solve()
x_values = {k: v.solution_value() for k, v in x.items()}

# === 3. AŞAMA: Satış kararlarını hesapla (y = min(x, z)) ===
satilanlar = {k: {} for k in range(SIMULASYON_SAYISI)}
for k in range(SIMULASYON_SAYISI):
    for u in urunler:
        uretim_miktari = sum(x_values[(u, j)] for j in ureticiler if (u, j) in x_values)
        talep = sales_scenarios[u][k]
        satilanlar[k][u] = min(uretim_miktari, talep)

# === 4. AŞAMA: Kar hesapla ===
karlar = []
uretim_maliyet = sum(x_values[(u, j)] * urun_uretici_dict[(u, j)] for (u, j) in x_values)
for k in range(SIMULASYON_SAYISI):
    gelir = sum(satilanlar[k][u] * satis_fiyat[u] for u in urunler)
    net_kar = gelir - uretim_maliyet
    karlar.append(net_kar)

ortalama_kar = np.mean(karlar)

# === Sonuçları Yazdır ===
print("=== Üretim Kararları (x) ===")
for (u, j), val in x_values.items():
    if val > 0:
        print(f"{u} - {j}: {val:.2f}")

print("\n=== Ortalama Kar (Senaryolar üzerinden): {:.2f} ===".format(ortalama_kar))
