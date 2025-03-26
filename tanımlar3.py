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
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")  # Üretici kapasite sayfası eklendi
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

# Gerekli tanımlamalar
urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))

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
    print(f"{i+1}. iterasyon için stokastik satış miktarları:", sales_stochastic)

    # Optimizasyon Modeli Kurulumu
    solver = pywraplp.Solver.CreateSolver('SCIP')
    
    # Değişken tanımları (Sürekli değişken olarak tanımlandı)
    x = {}
    uretim_kar = {}
    
    # Kar marjlarını hesapla (Fiyat - Maliyet)
    for urun in urunler:
        # Her ürün için üretim kararını belirlemek
        urun_kar = {}
        for uretici in ureticiler:
            fiyat = satis_fiyat[urun]
            maliyet = urun_uretici_dict.get((urun, uretici), None)
            
            if maliyet is not None:  # Eğer geçerli ikili varsa
                kar = fiyat - maliyet
                urun_kar[uretici] = kar
        uretim_kar[urun] = urun_kar
    
    # Ürünleri kar marjlarına göre sıralama
    urunler_sirala = sorted(urunler, key=lambda urun: max(uretim_kar[urun].values()) if urun_kar else 0, reverse=True)
    
    # En kârlı ürünleri üretim için seçme
    selected_urunler = urunler_sirala[:len(urunler)//2]  # Örnek olarak, en kârlı yarım ürünü seçebiliriz
    
    # Üretim kararlarını yalnızca seçilen ürünler için yap
    for urun in selected_urunler:
        urun_alt = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
        urun_ust = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
        
        for uretici in ureticiler:
            if (urun, uretici) in urun_uretici_dict:  # Eğer geçerli bir üretici-ürün ikilisi varsa
                # Üretim değişkeni oluştur
                x[(urun, uretici)] = solver.IntVar(0, urun_ust, f'x_{urun}_{uretici}')
    
    # Amaç Fonksiyonu
    objective = solver.Objective()
    for (urun, uretici), var in x.items():
        fiyat = satis_fiyat[urun]
        maliyet = urun_uretici_dict[(urun, uretici)]
        kar = fiyat - maliyet
        objective.SetCoefficient(var, kar)
    objective.SetMaximization()
    
    # Ürün Üretim Toplam Kısıtları
    for urun in selected_urunler:
        urun_toplam_kisit = solver.Add(
            sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x) >= 
            urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
        )
        urun_toplam_ust_kisit = solver.Add(
            sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x) <= 
            urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
        )
    
    # Modeli Çöz
    status = solver.Solve()
    
    if status == pywraplp.Solver.OPTIMAL:
        x_values = {key: var.solution_value() for key, var in x.items()}
        toplam_kar = sum(x_values[(urun, uretici)] * ((satis_fiyat[urun]) - urun_uretici_dict[(urun, uretici)]) for (urun, uretici) in x.keys())
        iterasyon_sonuclarini_yazdir(i, sales_stochastic, x_values, toplam_kar)
        SIMULASYON_SONUCLARI.append(toplam_kar)
    else:
        print(f"Optimizasyon problemi {i+1}. iterasyonda çözülemedi!")
        SIMULASYON_SONUCLARI.append(0)

# Simülasyon Sonuçlarının Analizi
simulasyon_df = pd.DataFrame(SIMULASYON_SONUCLARI, columns=["Toplam Kar"])

