import pandas as pd
from ortools.linear_solver import pywraplp

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



# Excel dosyasını yükleme (örneğin, 'satislar.xlsx')


# Sütun adlarını düzenleme
df_satis.columns = df_satis.columns.astype(str).str.replace(',', '').str.strip()

# Düzenlenmiş sütun adlarını kontrol etme
print("Düzenlenmiş sütun adları:", df_satis.columns)

# Yıl sütunlarını otomatik algılama
yillar = [col for col in df_satis.columns if col.isdigit()]
print("Yıl sütunları:", yillar)

# Ürün ortalama satış oranlarını hesaplama
urun_ortalama_satis = {}
for index, row in df_satis.iterrows():
    urun = row['Ürün']
    satis_oranlari = row[yillar]
    urun_ortalama_satis[urun] = satis_oranlari.mean()

# Ortalama satış oranlarını yazdırma
print("Ürün ortalama satış oranları:", urun_ortalama_satis)


maliyet_ust_siniri = df_kisitlar[df_kisitlar['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0]

# "Ürün - Satış" verilerini işleme
urun_ortalama_satis = {
    row['Ürün']: df_satis.loc[df_satis['Ürün'] == row['Ürün'], ['2020', '2021', '2022', '2023', '2024']]
    .replace('-', pd.NA).dropna(axis=1).mean(axis=1).iloc[0]
    for _, row in df_satis.iterrows()
}

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
