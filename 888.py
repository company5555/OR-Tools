import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp

# === Parametreler ===
SIMULASYON_SAYISI = 15
np.random.seed(12)
file_path = "ORTEST.xlsx"

# === Veri Okuma ===
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



# === Tüm Senaryoları Üret ===
sales_scenarios = []
for i in range(SIMULASYON_SAYISI):
    np.random.seed(12 + i)
    scenario = {
        urun: max(0, np.random.normal(urun_param_dict[urun]['ortalama'], urun_param_dict[urun]['std']))
        for urun in urunler
    }
    sales_scenarios.append(scenario)

# === Optimizasyon Modeli ===
solver = pywraplp.Solver.CreateSolver("SCIP")

# Karar Değişkenleri
x = {}
for urun in urunler:
    for uretici in ureticiler:
        if (urun, uretici) in urun_uretici_dict:
            x[(urun, uretici)] = solver.IntVar(0, solver.infinity(), f"x_{urun}_{uretici}")

# Senaryo Bazlı Talepleri Karşılayacak Üst Sınır Kısıtları
for s_index, scenario in enumerate(sales_scenarios):
    for urun in urunler:
        urun_toplam = sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x)
        uretilen_maks = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
        solver.Add(urun_toplam <= min(scenario[urun], uretilen_maks), f"senaryo_{s_index}_urun_{urun}")

# Üretici Kapasite Kısıtları
for uretici in ureticiler:
    toplam = sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x)
    solver.Add(toplam <= uretici_kapasite_dict.get(uretici, float("inf")), f"kapasite_{uretici}")
    solver.Add(toplam >= uretici_alt_kapasite_dict.get(uretici, 0), f"alt_kapasite_{uretici}")

# === Amaç Fonksiyonu: Ortalama Kar Maksimizasyonu ===
objective = solver.Objective()
for (urun, uretici), var in x.items():
    ortalama_kar = 0
    for scenario in sales_scenarios:
        satis_miktari = scenario[urun]
        kar = satis_fiyat[urun] - urun_uretici_dict[(urun, uretici)]
        ortalama_kar += kar / SIMULASYON_SAYISI  # Beklenen kâr
    objective.SetCoefficient(var, ortalama_kar)
objective.SetMaximization()

# === Modeli Çöz ===
status = solver.Solve()

# === Sonuçlar ===
if status == pywraplp.Solver.OPTIMAL:
    print("\n=== ORTAK MODEL SONUÇLARI ===")
    toplam_kar = objective.Value()
    for (urun, uretici), var in x.items():
        if var.solution_value() > 0:
            print(f"{urun} - {uretici}: {var.solution_value():.2f}")
    print(f"\nBeklenen Ortalama Kar: {toplam_kar:,.2f}")
else:
    print("Model optimal çözüme ulaşamadı.")


def iterasyon_sonuclarini_yazdir(iterasyon, sales_stochastic, x_values, toplam_kar, toplam_maliyet, uretici_toplam_uretim):
    print(f"\n=== {iterasyon + 1}. İTERASYON SONUÇLARI ===")
    
    # Stokastik Taleplerin Yazdırılması
    print("Stokastik Talepler (Satış Miktarları):")
    for urun, miktar in sales_stochastic.items():
        print(f"{urun}: {miktar:.2f}")
    
    # Üretim Miktarlarının Yazdırılması
    print("\nÜretim Miktarları:")
    for (urun, uretici), uretim_miktari in x_values.items():
        if uretim_miktari > 0:
            print(f"{urun} - {uretici}: {uretim_miktari:.2f}")

    # Toplam Üretim Maliyeti ve Kar
    print(f"\nToplam Üretim Maliyeti: {toplam_maliyet:,.2f}")
    print(f"Toplam Kar: {toplam_kar:,.2f}")

    # Üretici Bazında Toplam Üretim Miktarları
    print("\nÜretici Bazında Toplam Üretim Miktarları:")
    for uretici, miktar in uretici_toplam_uretim.items():
        print(f"{uretici}: {miktar:.2f}")