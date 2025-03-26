import numpy as np
import pandas as pd
from ortools.linear_solver import pywraplp

# Monte Carlo Simülasyonu Parametreleri
ITERASYON_SAYISI = 1000  # Simülasyon tekrar sayısı
SIMULASYON_SONUCLARI = []  # Sonuçları saklayacağımız liste

np.random.seed(42)  # Tekrar üretilebilir sonuçlar için

file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Satış")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")

# "Toplam Maliyet" Satırını Hariç Tutma ve Ürün İsimlerini Normalize Etme
urun_kisit_data = urun_kisit_data[urun_kisit_data['Ürün'] != "Toplam Maliyet"]

urunler = [urun for urun in urun_kisit_data['Ürün']]
# "Ürün - Param" sayfasından her ürün için ortalama ve standart sapma değerlerini alalım
urun_param_df = pd.read_excel("ORTEST.xlsx", sheet_name="Ürün - Param")
urun_param_dict = {
    row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]}
    for _, row in urun_param_df.iterrows()
}

for i in range(ITERASYON_SAYISI):
    # Stokastik Satış Miktarları (Normal Dağılım)
    sales_stochastic = {
        urun: max(0, np.random.normal(deger["ortalama"], deger["std"]))
        for urun, deger in urun_param_dict.items()
    }

    # Optimizasyon Modeli Kurulumu
    solver = pywraplp.Solver.CreateSolver('SCIP')

    x = { (urun, uretici): solver.IntVar(0, solver.infinity(), f'x_{urun}_{uretici}')
         for (urun, uretici) in urun_uretici_dict.keys() }

    y = { urun: solver.BoolVar(f'y_{urun}') for urun in urunler }
    z = { uretici: solver.BoolVar(f'z_{uretici}') for uretici in ureticiler }

    # Üretim Kısıtları
    for urun in urunler:
        urun_alt = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
        urun_ust = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
        gecerli_uretici_ciftleri = [uretici for uretici in ureticiler if (urun, uretici) in x]

        solver.Add(solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) >= urun_alt * y[urun])
        solver.Add(solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) <= urun_ust * y[urun])

    # Toplam Maliyet Kısıtı
    toplam_maliyet_ust_sinir = float(toplam_maliyet_data[toplam_maliyet_data['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0])
    solver.Add(solver.Sum(x[(urun, uretici)] * urun_uretici_dict[(urun, uretici)] for (urun, uretici) in x.keys()) <= toplam_maliyet_ust_sinir)

    # Üretici Kapasite Kısıtları
    for uretici in ureticiler:
        gecerli_urun_ciftleri = [urun for urun in urunler if (urun, uretici) in x]
        alt_kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Alt Kapasite'].values[0]
        ust_kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Üst Kapasite'].values[0]

        solver.Add(solver.Sum(x[(urun, uretici)] for urun in gecerli_urun_ciftleri) >= alt_kapasite * z[uretici])
        solver.Add(solver.Sum(x[(urun, uretici)] for urun in gecerli_urun_ciftleri) <= ust_kapasite * z[uretici])
        for urun in gecerli_urun_ciftleri:
            solver.Add(x[(urun, uretici)] <= ust_kapasite * z[uretici])

    # Amaç Fonksiyonu (Stokastik Kar Maksimizasyonu)
    objective = solver.Objective()
    for (urun, uretici), var in x.items():
        fiyat = satis_fiyat[urun]
        satis_miktari = sales_stochastic[urun]  # Rastgele çekilen satış miktarı
        maliyet = urun_uretici_dict[(urun, uretici)]
        net_kar = (fiyat * satis_miktari) - maliyet
        objective.SetCoefficient(var, net_kar)

    objective.SetMaximization()

    # Modeli Çöz
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        toplam_kar = sum(
            x[(urun, uretici)].solution_value() * ((satis_fiyat[urun] * sales_stochastic[urun]) - urun_uretici_dict[(urun, uretici)])
            for (urun, uretici) in x.keys()
        )

        SIMULASYON_SONUCLARI.append(toplam_kar)

# Simülasyon Sonuçlarının Analizi
simulasyon_df = pd.DataFrame(SIMULASYON_SONUCLARI, columns=["Toplam Kar"])
ortalama_kar = simulasyon_df["Toplam Kar"].mean()
std_kar = simulasyon_df["Toplam Kar"].std()
min_kar = simulasyon_df["Toplam Kar"].min()
max_kar = simulasyon_df["Toplam Kar"].max()
q1, q3 = simulasyon_df["Toplam Kar"].quantile([0.25, 0.75])

print("\n=== MONTE CARLO SİMÜLASYONU SONUÇLARI ===")
print(f"Ortalama Kar: {ortalama_kar:,.2f}")
print(f"Standart Sapma: {std_kar:,.2f}")
print(f"Min Kar: {min_kar:,.2f}")
print(f"Max Kar: {max_kar:,.2f}")
print(f"1. Çeyrek (Q1): {q1:,.2f}")
print(f"3. Çeyrek (Q3): {q3:,.2f}")
