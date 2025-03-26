import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp

# Excel dosyasını okuma
file_path = "ORTEST.xlsx"

urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Satış")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")

# Verileri ön işleme
urun_kisit_data = urun_kisit_data[urun_kisit_data['Ürün'] != "Toplam Maliyet"]

urunler = urun_kisit_data['Ürün'].tolist()

# Satış verilerini normal dağılım için hazırla
sales_data = urun_satis_data.set_index('Ürün').iloc[:, 1:].values
sales_mean = sales_data.mean(axis=1)
sales_stddev = sales_data.std(axis=1)

# Ürün ve üretici maliyet verilerini sözlük olarak hazırla
urun_uretici_dict = {
    (row['Ürün'], row['Üretici']): row['Birim Maliyet']
    for _, row in urun_uretici_data.iterrows()
}

# Ürünlerin satış fiyatlarını hazırlama
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))

# Solver ve değişkenler
solver = pywraplp.Solver.CreateSolver('SCIP')

# Üretim miktarı değişkenleri
x = {}
for urun in urunler:
    for uretici in uretici_kapasite_data['Üretici'].unique():
        x[(urun, uretici)] = solver.IntVar(0, solver.infinity(), f'x_{urun}_{uretici}')

# Üretim kararı binary değişkeni
y = {urun: solver.BoolVar(f'y_{urun}') for urun in urunler}

# Üretici kullanım kararı binary değişkeni
z = {uretici: solver.BoolVar(f'z_{uretici}') for uretici in uretici_kapasite_data['Üretici'].unique()}

# Kısıtlar
for urun in urunler:
    urun_alt = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
    urun_ust = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
    gecerli_uretici_ciftleri = [(urun, uretici) for uretici in uretici_kapasite_data['Üretici'].unique()]
    
    solver.Add(solver.Sum(x[(urun, uretici)] for urun, uretici in gecerli_uretici_ciftleri) >= urun_alt * y[urun])
    solver.Add(solver.Sum(x[(urun, uretici)] for urun, uretici in gecerli_uretici_ciftleri) <= urun_ust * y[urun])

# Toplam maliyet kısıtı
toplam_maliyet_ust_sinir = float(toplam_maliyet_data[toplam_maliyet_data['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0])
solver.Add(solver.Sum(x[(urun, uretici)] * urun_uretici_dict.get((urun, uretici), 0) for urun in urunler for uretici in uretici_kapasite_data['Üretici'].unique()) <= toplam_maliyet_ust_sinir)

# Üretici kapasite kısıtları
for uretici in uretici_kapasite_data['Üretici'].unique():
    gecerli_urun_ciftleri = [(urun, uretici) for urun in urunler]
    alt_kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Alt Kapasite'].values[0]
    ust_kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Üst Kapasite'].values[0]
    
    solver.Add(solver.Sum(x[(urun, uretici)] for urun, uretici in gecerli_urun_ciftleri) >= alt_kapasite * z[uretici])
    solver.Add(solver.Sum(x[(urun, uretici)] for urun, uretici in gecerli_urun_ciftleri) <= ust_kapasite * z[uretici])

# Amaç fonksiyonu - Stokastik
num_simulations = 1000
objective = solver.Objective()

for urun in urunler:
    for uretici in uretici_kapasite_data['Üretici'].unique():
        mean_sales = sales_mean[urun]
        stddev_sales = sales_stddev[urun]
        
        # Her simülasyon için satışları rastgele çek
        simulated_sales = np.random.normal(mean_sales, stddev_sales, num_simulations)
        
        # Beklenen karın hesaplanması
        expected_profit = np.mean([(sim_sales * satis_fiyat[urun] * 0.5) - urun_uretici_dict.get((urun, uretici), 0) for sim_sales in simulated_sales])
        objective.SetCoefficient(x[(urun, uretici)], expected_profit)

objective.SetMaximization()

# Çözümü bulma
status = solver.Solve()

# Sonuçları yazdırma
if status == pywraplp.Solver.OPTIMAL:
    print("Optimal çözüm bulundu!\n")
    
    toplam_uretim = 0
    toplam_maliyet = 0
    toplam_beklenen_gelir = 0
    
    # Üretim kararları
    for urun in urunler:
        for uretici in uretici_kapasite_data['Üretici'].unique():
            if x[(urun, uretici)].solution_value() > 0:
                print(f"{urun} için {uretici} üretildi, Adet: {x[(urun, uretici)].solution_value()}")
    
    print("\n=== GENEL ÖZET ===")
    print(f"Toplam Üretim Adedi: {toplam_uretim}")
    print(f"Toplam Maliyet: {toplam_maliyet}")
    print(f"Toplam Beklenen Gelir: {toplam_beklenen_gelir}")
else:
    print("Optimal çözüm bulunamadı.")
