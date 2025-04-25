import pandas as pd
import numpy as np
import time
from ortools.linear_solver import pywraplp
from tqdm import tqdm

# Parametreler
SIMULASYON_SAYISI = 25000
np.random.seed(25)

# Excel'den veri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

# Ürün ve üretici listeleri
urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))

# Verileri sözlüklere çevirme
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# Talep senaryoları üret
sales_scenarios = {
    urun: [max(0, int(round(np.random.normal(p["ortalama"], p["std"]))))
           for _ in range(SIMULASYON_SAYISI)]
    for urun, p in urun_param_dict.items()
}

# Optimizasyon modeli başlat
start_time = time.time()
solver = pywraplp.Solver.CreateSolver('SCIP')

# Karar değişkenleri: her ürün-üretici için sabit üretim kararı
x = {
    (urun, uretici): solver.IntVar(0, sum(sales_scenarios[urun]) // SIMULASYON_SAYISI + 100,
                                   f'x_{urun}_{uretici}')
    for urun in urunler for uretici in ureticiler if (urun, uretici) in urun_uretici_dict
}

# Yardımcı değişkenler: senaryolara göre satılabilen miktar
satilan = {
    (urun, s): solver.NumVar(0, solver.infinity(), f'satilan_{urun}_{s}')
    for urun in urunler for s in range(SIMULASYON_SAYISI)
}

# Satış kısıtları
for urun in urunler:
    toplam_uretim = sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x)
    for s in range(SIMULASYON_SAYISI):
        solver.Add(satilan[(urun, s)] <= toplam_uretim)
        solver.Add(satilan[(urun, s)] <= sales_scenarios[urun][s])

# Üretici kapasite kısıtları
for uretici in ureticiler:
    toplam = sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x)
    solver.Add(toplam <= uretici_kapasite_dict.get(uretici, float('inf')))
    solver.Add(toplam >= uretici_alt_kapasite_dict.get(uretici, 0))

# Amaç fonksiyonu: Ortalama karı maksimize et (progress bar ile)
objective = solver.Objective()
progress_bar = tqdm(total=len(urunler) * SIMULASYON_SAYISI, desc="Amaç Fonksiyonu Hazırlanıyor", unit="ürün-senaryo")
start_obj_time = time.time()

for urun in urunler:
    for s in range(SIMULASYON_SAYISI):
        objective.SetCoefficient(satilan[(urun, s)], satis_fiyat[urun] / SIMULASYON_SAYISI)
        progress_bar.update(1)
        elapsed = time.time() - start_obj_time
        avg_time = elapsed / progress_bar.n if progress_bar.n else 0
        remaining = avg_time * (progress_bar.total - progress_bar.n)
        progress_bar.set_postfix_str(f"Kalan süre: {remaining:.1f}s")

for (urun, uretici), var in x.items():
    objective.SetCoefficient(var, -urun_uretici_dict[(urun, uretici)])

progress_bar.close()
objective.SetMaximization()

# Modeli çöz
if solver.Solve() == pywraplp.Solver.OPTIMAL:
    print("\n=== Sonuçlar ===")
    toplam_kar = 0
    toplam_gider = 0
    for urun in urunler:
        uretim = sum(x[(urun, uretici)].solution_value() for uretici in ureticiler if (urun, uretici) in x)
        satilanlar = [satilan[(urun, s)].solution_value() for s in range(SIMULASYON_SAYISI)]       
        toplam_kar += sum(satilanlar) * satis_fiyat[urun] / SIMULASYON_SAYISI
    print("\n=== Ürün Bazında Toplam Üretim Miktarları ===")
    for urun in urunler:
        toplam_uretim = sum(x[(urun, uretici)].solution_value() for uretici in ureticiler if (urun, uretici) in x)
        print(f"{urun}: {toplam_uretim:.0f} adet")


    for (urun, uretici), var in x.items(): 
        toplam_gider += var.solution_value() * urun_uretici_dict[(urun, uretici)]
    print(f"\nToplam Gider: {toplam_gider:,.2f}")
    print(f"Beklenen Ortalama Kar: {toplam_kar - toplam_gider:,.2f}")
    print(f"Çözüm süresi: {time.time() - start_time:.2f} saniye")
else:
    print("Optimizasyon çözülemedi.")
