import pandas as pd
from ortools.linear_solver import pywraplp
import time

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
talep_dict = dict(zip(urun_param_df['Ürün'], urun_param_df['Ortalama']))
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# Optimizasyon modeli başlat
solver = pywraplp.Solver.CreateSolver('SCIP')
start_time = time.time()

# Karar değişkenleri
x = {
    (urun, uretici): solver.IntVar(0, talep_dict[urun] + 100, f'x_{urun}_{uretici}')
    for urun in urunler for uretici in ureticiler if (urun, uretici) in urun_uretici_dict
}

# Kısıtlar: Talebi aşma
for urun in urunler:
    toplam_uretim = sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x)
    solver.Add(toplam_uretim <= talep_dict[urun])

# Üretici kapasite kısıtları
for uretici in ureticiler:
    toplam = sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x)
    solver.Add(toplam <= uretici_kapasite_dict.get(uretici, float('inf')))
    solver.Add(toplam >= uretici_alt_kapasite_dict.get(uretici, 0))

# Amaç fonksiyonu: Karı maksimize et
objective = solver.Objective()
for (urun, uretici), var in x.items():
    gelir = satis_fiyat[urun]
    maliyet = urun_uretici_dict[(urun, uretici)]
    objective.SetCoefficient(var, gelir - maliyet)

objective.SetMaximization()

# Modeli çöz
if solver.Solve() == pywraplp.Solver.OPTIMAL:
    print("\n=== SONUÇLAR ===")
    toplam_kar = 0
    toplam_gider = 0

    print("\n=== Ürün Bazında Üretim Miktarları ===")
    for urun in urunler:
        urun_uretim = sum(x[(urun, uretici)].solution_value() for uretici in ureticiler if (urun, uretici) in x)
        print(f"{urun}: {urun_uretim:.0f} adet")
        toplam_kar += urun_uretim * satis_fiyat[urun]

    for (urun, uretici), var in x.items():
        toplam_gider += var.solution_value() * urun_uretici_dict[(urun, uretici)]

    print(f"\nToplam Gider: {toplam_gider:,.2f}")
    print(f"Toplam Kar: {toplam_kar - toplam_gider:,.2f}")
    print(f"Çözüm Süresi: {time.time() - start_time:.2f} saniye")

    print("\n=== Üretici Bazında Üretim Miktarları ===")
    for uretici in ureticiler:
        uretici_uretim = sum(x[(urun, uretici)].solution_value() for urun in urunler if (urun, uretici) in x)
        print(f"{uretici}: {uretici_uretim:.0f} adet")
else:
    print("Optimizasyon problemi çözülemedi.")
