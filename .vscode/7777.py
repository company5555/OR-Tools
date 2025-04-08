import pandas as pd
import numpy as np
import time
from ortools.linear_solver import pywraplp

# Monte Carlo Simülasyonu Parametreleri
SIMULASYON_SAYISI = 10
np.random.seed(12)

# Excel dosyasından verileri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
# Ürün ve Üretici Listeleri
urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))

# Sözlükler
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# Stokastik satış tahminlerini oluşturma
sales_scenarios = {urun: [max(0, int(round(np.random.normal(deger["ortalama"], deger["std"])))) for _ in range(SIMULASYON_SAYISI)] for urun, deger in urun_param_dict.items()}

# Optimizasyon başlat
start_time = time.time()
solver = pywraplp.Solver.CreateSolver('SCIP')
x = {(urun, uretici): solver.IntVar(0, max(sales_scenarios[urun]), f'x_{urun}_{uretici}') for urun in urunler for uretici in ureticiler if (urun, uretici) in urun_uretici_dict}

# Amaç fonksiyonu: Beklenen karı maksimize etme
objective = solver.Objective()
for (urun, uretici), var in x.items():
    kar = sum((satis_fiyat[urun] - urun_uretici_dict[(urun, uretici)]) * (talep / SIMULASYON_SAYISI) for talep in sales_scenarios[urun])
    objective.SetCoefficient(var, kar)
objective.SetMaximization()

# Üretim kapasite kısıtları
for urun in urunler:
    solver.Add(sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x) <= max(sales_scenarios[urun]))

for uretici in ureticiler:
    solver.Add(sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x) <= uretici_kapasite_dict.get(uretici, float('inf')))
    solver.Add(sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x) >= uretici_alt_kapasite_dict.get(uretici, 0))

# Optimizasyonu çöz
if solver.Solve() == pywraplp.Solver.OPTIMAL:
    toplam_kar = sum(x[(urun, uretici)].solution_value() * (satis_fiyat[urun] - urun_uretici_dict[(urun, uretici)]) for (urun, uretici) in x.keys())
    toplam_gider = sum(x[(urun, uretici)].solution_value() * urun_uretici_dict[(urun, uretici)] for (urun, uretici) in x.keys())
    
    print("\n=== Simülasyon Sonuçları ===")
    for urun in urunler:
        print(f"{urun} için Talep Senaryoları: {sales_scenarios[urun]}")
        uretilen_miktar = sum(x[(urun, uretici)].solution_value() for uretici in ureticiler if (urun, uretici) in x)
        print(f"Modelin Karar Verdiği Üretim: {uretilen_miktar}")
    
    print("\nÜretici Bazında Üretim Miktarları:")
    for uretici in ureticiler:
        uretim_miktari = sum(x[(urun, uretici)].solution_value() for urun in urunler if (urun, uretici) in x)
        print(f"{uretici}: {uretim_miktari}")
    
    print(f"\nToplam Gider: {toplam_gider:,.2f}")
    print(f"Beklenen Ortalama Kar: {toplam_kar:,.2f}")
    execution_time = time.time() - start_time
    print(f"İşlenme süresi: {execution_time:.4f} saniye")
else:
    print("Optimizasyon problemi çözülemedi!")
