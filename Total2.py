import pandas as pd
from ortools.linear_solver import pywraplp

# Uyarıları kapatma
pd.set_option('mode.chained_assignment', None)

# Verileri Excel'den okuma
file_path = "ORTEST.xlsx"
df_kisitlar = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
df_satis = pd.read_excel(file_path, sheet_name="Ürün - Satış")
df_uretici = pd.read_excel(file_path, sheet_name="Ürün - Üretici")

# "Ürün - Kısıt" verilerini işleme
urun_kisitlari = {
    row['Ürün']: {
        'alt_sinir': row['Üretim Alt Sınır'],
        'ust_sinir': row['Üretim Üst Sınır']
    }
    for _, row in df_kisitlar.iterrows() if row['Ürün'] != 'Toplam Maliyet'
}

maliyet_ust_siniri = df_kisitlar[df_kisitlar['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0]

# "Ürün - Satış" verilerini işleme
urun_ortalama_satis = {}
for _, row in df_satis.iterrows():
    urun = row['Ürün']
    satis_verileri = row[['A', 'B', 'C', 'D', 'E']].replace('-', pd.NA)
    satis_verileri = satis_verileri.dropna().astype(float)  # Eksik değerleri kaldır ve sayıya çevir
    if not satis_verileri.empty:
        urun_ortalama_satis[urun] = satis_verileri.mean()
    else:
        urun_ortalama_satis[urun] = 0  # Eksik değerler varsa ortalama 0 olarak atanır

urun_satis_fiyatlari = {
    row['Ürün']: row['Satış Fiyatı'] for _, row in df_satis.iterrows()
}

# "Ürün - Üretici" verilerini işleme
ureticiler = {}
for _, row in df_uretici.iterrows():
    uretici = row['Üretici']
    kapasite = row['Kapasite']
    ureticiler[uretici] = {
        'kapasite': kapasite,
        'urunler': {
            row[f'Ürün {i}']: row[f'Birim Maliyet {i}']
            for i in range(1, len(row)//2) if pd.notna(row[f'Ürün {i}'])
        }
    }

# OR-Tools çözümleyicisini oluşturma
solver = pywraplp.Solver.CreateSolver('SCIP')
if not solver:
    raise Exception("Solver oluşturulamadı!")

# Karar değişkenleri
x = {}
for uretici, data in ureticiler.items():
    for urun, maliyet in data['urunler'].items():
        x[uretici, urun] = solver.IntVar(0, data['kapasite'], f'x_{uretici}_{urun}')

# Amaç fonksiyonu
objective = solver.Objective()
for (uretici, urun), var in x.items():
    if urun in urun_ortalama_satis:
        satis_ortalamasi = urun_ortalama_satis[urun]
        satis_fiyati = urun_satis_fiyatlari[urun]
        birim_maliyet = ureticiler[uretici]['urunler'][urun]
        objective.SetCoefficient(var, (satis_ortalamasi * satis_fiyati) - birim_maliyet)
objective.SetMaximization()

# Kısıtlar
# Ürün üretim sınırları
for urun, kisit in urun_kisitlari.items():
    urun_vars = [x[uretici, urun] for uretici in ureticiler if (uretici, urun) in x]
    if urun_vars:
        toplam_uretim = solver.Sum(urun_vars)
        if not pd.isna(kisit['alt_sinir']):
            solver.Add(toplam_uretim >= kisit['alt_sinir'])
        if not pd.isna(kisit['ust_sinir']):
            solver.Add(toplam_uretim <= kisit['ust_sinir'])

# Üretici kapasite sınırları
for uretici, data in ureticiler.items():
    toplam_kullanim = solver.Sum(x[uretici, urun] for urun in data['urunler'] if (uretici, urun) in x)
    solver.Add(toplam_kullanim <= data['kapasite'])

# Toplam maliyet sınırı
toplam_maliyet = solver.Sum(
    x[uretici, urun] * ureticiler[uretici]['urunler'][urun]
    for (uretici, urun) in x
)
solver.Add(toplam_maliyet <= maliyet_ust_siniri)

# Çözümü bulma
status = solver.Solve()

# Çıktılar
if status == pywraplp.Solver.OPTIMAL:
    print("SATIŞ OLASILIKLARI:")
    for urun, satis_ortalamasi in urun_ortalama_satis.items():
        print(f"  {urun}: {satis_ortalamasi:.2f}")

    print("\nÜRETİM SONUÇLARI:")
    z_degeri = 0
    for (uretici, urun), var in x.items():
        if var.solution_value() > 0:
            satis_ortalamasi = urun_ortalama_satis[urun]
            satis_fiyati = urun_satis_fiyatlari[urun]
            birim_maliyet = ureticiler[uretici]['urunler'][urun]
            kar = (satis_ortalamasi * satis_fiyati - birim_maliyet) * var.solution_value()
            z_degeri += kar
            print(f"  Üretici: {uretici}, Ürün: {urun}, Üretim: {var.solution_value()}, Kar: {kar:.2f}")

    print(f"\nZ Değeri (Amaç Fonksiyonu): {z_degeri:.2f}")
else:
    print("Çözüm bulunamadı.")


# Çözüm öncesi kısıt ve değişken detayları
print("=== Ürün Üretim Sınırları ===")
for urun, kisit in urun_kisitlari.items():
    print(f"{urun}: Alt Sınır = {kisit['alt_sinir']}, Üst Sınır = {kisit['ust_sinir']}")

print("\n=== Üretici Kapasiteleri ===")
for uretici, data in ureticiler.items():
    print(f"{uretici}: Kapasite = {data['kapasite']}")

print("\n=== Toplam Maliyet Üst Sınırı ===")
print(f"Maliyet Üst Sınırı: {maliyet_ust_siniri}")

print("\n=== Satış Ortalama ve Fiyatları ===")
for urun, ortalama in urun_ortalama_satis.items():
    fiyat = urun_satis_fiyatlari.get(urun, 0)
    print(f"{urun}: Ortalama = {ortalama}, Fiyat = {fiyat}")

# Çözüm süreci
status = solver.Solve()

# Çözüm durumu analizi
if status == pywraplp.Solver.OPTIMAL:
    print("Optimal çözüm bulundu!")
elif status == pywraplp.Solver.FEASIBLE:
    print("Feasible (uygun) bir çözüm bulundu!")
elif status == pywraplp.Solver.INFEASIBLE:
    print("Çözüm bulunamadı! Model infeasible (uygunsuz).")
elif status == pywraplp.Solver.UNBOUNDED:
    print("Çözüm bulunamadı! Model unbounded (sınırsız).")
else:
    print("Çözüm bulunamadı! Sebep bilinmiyor.")

