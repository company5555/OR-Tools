import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp

# Monte Carlo Simülasyonu Parametreleri
ITERASYON_SAYISI = 1
SIMULASYON_SONUCLARI = []

np.random.seed(51)

# Excel dosyasından verileri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

# Gerekli tanımlamalar
urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = [uretici for uretici in uretici_kapasite_data['Üretici']]

# Sözlük hazırlama
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}

def iterasyon_sonuclarini_yazdir(iterasyon, sales_stochastic, x_values, toplam_kar):
    print(f"\n=== {iterasyon + 1}. İTERASYON SONUÇLARI ===")
    print("Stokastik Satış Miktarları:")
    for urun, miktar in sales_stochastic.items():
        print(f"{urun}: {miktar:.2f}")
    
    print("\nÜretim Miktarları:")
    for (urun, uretici), uretim_miktari in x_values.items():
        if uretim_miktari > 0:
            print(f"{urun} - {uretici}: {uretim_miktari:.2f}")
    
    print(f"\nToplam Kar: {toplam_kar:,.2f}")

# Monte Carlo Simülasyonu
for i in range(ITERASYON_SAYISI):
    # Stokastik Satış Miktarları (Normal Dağılım)
    sales_stochastic = {
        urun: max(0, np.random.normal(deger["ortalama"], deger["std"]))
        for urun, deger in urun_param_dict.items()
    }

    # Optimizasyon Modeli Kurulumu
    solver = pywraplp.Solver.CreateSolver('SCIP')

    # Her ürün için üretim sınırlarını hesapla
    urun_uretim_sinirlar = {}
    for urun in urunler:
        urun_alt = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
        urun_ust = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
        
        # Alt sınır için max{ürün üretim alt sınırı, satış miktarı}
        alt_sinir = max(urun_alt, sales_stochastic[urun])
        
        # Üst sınır için min{Üretim üst sınırı, belirlenen satış miktarı}
        ust_sinir = min(urun_ust, sales_stochastic[urun])
        
        urun_uretim_sinirlar[urun] = {
            'alt_sinir': alt_sinir,
            'ust_sinir': ust_sinir
        }

    # Değişken tanımları
    x = { (urun, uretici): solver.IntVar(0, urun_uretim_sinirlar[urun]['ust_sinir'], f'x_{urun}_{uretici}')
          for (urun, uretici) in urun_uretici_dict.keys() }

    y = { urun: solver.BoolVar(f'y_{urun}') for urun in urunler }
    z = { uretici: solver.BoolVar(f'z_{uretici}') for uretici in ureticiler }

    # Üretim Kısıtları
    for urun in urunler:
        gecerli_uretici_ciftleri = [uretici for uretici in ureticiler if (urun, uretici) in x]
        alt_sinir = urun_uretim_sinirlar[urun]['alt_sinir']
        ust_sinir = urun_uretim_sinirlar[urun]['ust_sinir']

        solver.Add(solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) >= alt_sinir * y[urun])
        solver.Add(solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) <= ust_sinir * y[urun])

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
        maliyet = urun_uretici_dict[(urun, uretici)]
        objective.SetCoefficient(var, (fiyat - maliyet))

    objective.SetMaximization()

    # Modeli Çöz
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        # Çözüm değerlerini al
        x_values = {key: var.solution_value() for key, var in x.items()}
        
        toplam_kar = sum(
            x_values[(urun, uretici)] * ((satis_fiyat[urun]) - urun_uretici_dict[(urun, uretici)])
            for (urun, uretici) in x.keys()
        )

        # Her iterasyonun detaylı sonuçlarını yazdır
        iterasyon_sonuclarini_yazdir(i, sales_stochastic, x_values, toplam_kar)

        SIMULASYON_SONUCLARI.append(toplam_kar)
    print(gecerli_uretici_ciftleri)
# Simülasyon Sonuçlarının Analizi
simulasyon_df = pd.DataFrame(SIMULASYON_SONUCLARI, columns=["Toplam Kar"])
ortalama_kar = simulasyon_df["Toplam Kar"].mean()
std_kar = simulasyon_df["Toplam Kar"].std()
min_kar = simulasyon_df["Toplam Kar"].min()
max_kar = simulasyon_df["Toplam Kar"].max()
q1, q3 = simulasyon_df["Toplam Kar"].quantile([0.25, 0.75])

print("\n=== MONTE CARLO SİMÜLASYONU GENEL SONUÇLARI ===")
print(f"Ortalama Kar: {ortalama_kar:,.2f}")
print(f"Standart Sapma: {std_kar:,.2f}")
print(f"Min Kar: {min_kar:,.2f}")
print(f"Max Kar: {max_kar:,.2f}")
print(f"1. Çeyrek (Q1): {q1:,.2f}")
print(f"3. Çeyrek (Q3): {q3:,.2f}")




