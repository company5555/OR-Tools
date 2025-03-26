import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp

# Monte Carlo Simülasyonu Parametreleri
ITERASYON_SAYISI = 1
SIMULASYON_SONUCLARI = []

np.random.seed(12)

# Excel dosyasından verileri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

# Ürün ve Üretici Listeleri
urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))

# Sözlükler
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}


# Üretici Kapasite Verilerini Sözlük Haline Getirme
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))


def iterasyon_sonuclarini_yazdir(iterasyon, sales_stochastic, x_values, toplam_kar, toplam_maliyet, uretici_toplam_uretim):
    print(f"\n=== {iterasyon + 1}. İTERASYON SONUÇLARI ===")
    
    print("Stokastik Satış Miktarları:")
    for urun, miktar in sales_stochastic.items():
        print(f"{urun}: {miktar:.2f}")
    
    print("\nÜretim Miktarları:")
    for (urun, uretici), uretim_miktari in x_values.items():
        if uretim_miktari > 0:
            print(f"{urun} - {uretici}: {uretim_miktari:.2f}")

    print(f"\nToplam Üretim Maliyeti: {toplam_maliyet:,.2f}")
    print(f"Toplam Kar: {toplam_kar:,.2f}")

    print("\nÜretici Bazında Toplam Üretim Miktarları:")
    for uretici, miktar in uretici_toplam_uretim.items():
        print(f"{uretici}: {miktar:.2f}")
    
# Monte Carlo Simülasyonu
for i in range(ITERASYON_SAYISI):
    # Stokastik Satış Miktarları (Normal Dağılım)
    sales_stochastic = {urun: max(0, np.random.normal(deger["ortalama"], deger["std"])) for urun, deger in urun_param_dict.items()}
    
    # Ürün Üst Sınırlarını Belirleme
    urun_ust_kisit = {urun: min(sales_stochastic[urun], urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]) for urun in urunler}
    urun_alt_kisit = {urun: max(sales_stochastic[urun], urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]) for urun in urunler}
    # Optimizasyon Modeli Kurulumu
    solver = pywraplp.Solver.CreateSolver('SCIP')
    
    # Üretim Değişkenleri
    x = {}
    for urun in urunler:
        for uretici in ureticiler:
            if (urun, uretici) in urun_uretici_dict:
                x[(urun, uretici)] = solver.IntVar(0, urun_ust_kisit[urun], f'x_{urun}_{uretici}')
        
    # Amaç Fonksiyonu
    objective = solver.Objective()
    for (urun, uretici), var in x.items():
        kar = satis_fiyat[urun] - urun_uretici_dict[(urun, uretici)]
        objective.SetCoefficient(var, kar)
    objective.SetMaximization()
    
    # Ürün Üretim Toplam Kısıtları
    for urun in urunler:
        solver.Add(sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x) <= urun_ust_kisit[urun])
    
    # Üretici Kapasite Kısıtı
    for uretici in ureticiler:
        solver.Add(sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x) <= uretici_kapasite_dict.get(uretici, float('inf')))

    for uretici in ureticiler:
        solver.Add(sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x) >= uretici_alt_kapasite_dict.get(uretici, float('inf')))

    # Modeli Çöz
    status = solver.Solve()
    
    if status == pywraplp.Solver.OPTIMAL:
        x_values = {key: var.solution_value() for key, var in x.items()}
        
        # Toplam Üretim Maliyeti
        toplam_maliyet = sum(x_values[(urun, uretici)] * urun_uretici_dict[(urun, uretici)] for (urun, uretici) in x.keys())
        
        # Toplam Kar
        toplam_kar = sum(x_values[(urun, uretici)] * (satis_fiyat[urun] - urun_uretici_dict[(urun, uretici)]) for (urun, uretici) in x.keys())
        
        # Üretici Bazında Toplam Üretim Miktarları
        uretici_toplam_uretim = {uretici: sum(x_values[(urun, uretici)] for urun in urunler if (urun, uretici) in x_values) for uretici in ureticiler}

        # Sonuçları Yazdır
        iterasyon_sonuclarini_yazdir(i, sales_stochastic, x_values, toplam_kar, toplam_maliyet, uretici_toplam_uretim)
        
        # Sonuçları Kaydet
        SIMULASYON_SONUCLARI.append(toplam_kar)
    else:
        print(f"Optimizasyon problemi {i+1}. iterasyonda çözülemedi!")
        SIMULASYON_SONUCLARI.append(0)

# Simülasyon Sonuçlarının Analizi
simulasyon_df = pd.DataFrame(SIMULASYON_SONUCLARI, columns=["Toplam Kar"])

