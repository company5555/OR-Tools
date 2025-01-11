import pandas as pd
from ortools.linear_solver import pywraplp
pd.set_option('future.no_silent_downcasting', True)

# Verileri Yükleme
file_path = "ORTEST.xlsx"

urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Satış")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")

# "Toplam Maliyet" Satırını Hariç Tutma ve Ürün İsimlerini Normalize Etme
urun_kisit_data = urun_kisit_data[urun_kisit_data['Ürün'] != "Toplam Maliyet"]

# Ürünlerin tanımı
urunler = [urun for urun in urun_kisit_data['Ürün']]

sales_probability = dict(
    zip(
        urun_satis_data['Ürün'],
        urun_satis_data.iloc[:, 1:6].replace("-", 0).mean(axis=1)
    )
)

# Satış fiyatlarının tanımı
satis_fiyat = dict(
    zip(
        urun_satis_data['Ürün'],
        urun_satis_data['Satış Fiyatı']
    )
)

# Üreticilerin tanımlanması
ureticiler = [uretici for uretici in uretici_kapasite_data['Üretici']]

# Ürün ve Üretici Maliyet Bilgilerini Dictionary Olarak Hazırlama
urun_uretici_dict = {
    (row['Ürün'], row['Üretici']): row['Birim Maliyet']
    for _, row in urun_uretici_data.iterrows()
}

# Solver ve Değişkenlerin Tanımlanması
solver = pywraplp.Solver.CreateSolver('SCIP')

# Üretim miktarı değişkenleri
x = {}
for (urun, uretici) in urun_uretici_dict.keys():
    x[(urun, uretici)] = solver.IntVar(0, solver.infinity(), f'x_{urun}_{uretici}')

# Ürün üretim kararı için binary değişkenler (0: üretilmeyecek, 1: üretilecek)
y = {}
for urun in urunler:
    y[urun] = solver.BoolVar(f'y_{urun}')

# Üretici kullanım kararı için binary değişkenler (0: kullanılmayacak, 1: kullanılacak)
z = {}
for uretici in ureticiler:
    z[uretici] = solver.BoolVar(f'z_{uretici}')

# Kısıtlar
# Ürün Alt ve Üst Üretim Sınırı
for urun in urunler:
    urun_alt = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
    urun_ust = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
    
    gecerli_uretici_ciftleri = [uretici for uretici in ureticiler if (urun, uretici) in x]
    
    # Eğer ürün üretilecekse (y[urun] = 1), alt sınır kısıtı geçerli olur
    solver.Add(
        solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) >= urun_alt * y[urun]
    )
    # Üst sınır kısıtı her zaman geçerli
    solver.Add(
        solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) <= urun_ust * y[urun]
    )

# Toplam Maliyet Kısıtı
toplam_maliyet_ust_sinir = float(toplam_maliyet_data[toplam_maliyet_data['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0])

solver.Add(
    solver.Sum(
        x[(urun, uretici)] * urun_uretici_dict.get((urun, uretici), 0)
        for (urun, uretici) in x.keys()
    ) <= toplam_maliyet_ust_sinir
)

# Üretici Kapasite Kısıtları
for uretici in ureticiler:
    gecerli_urun_ciftleri = [urun for urun in urunler if (urun, uretici) in x]
    
    alt_kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Alt Kapasite'].values[0]
    ust_kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Üst Kapasite'].values[0]
    
    # Üretici kullanılıyorsa (z[uretici] = 1), alt kapasite kısıtı geçerli olur
    solver.Add(
        solver.Sum(x[(urun, uretici)] for urun in gecerli_urun_ciftleri) >= alt_kapasite * z[uretici]
    )
    # Üst kapasite kısıtı her zaman geçerli
    solver.Add(
        solver.Sum(x[(urun, uretici)] for urun in gecerli_urun_ciftleri) <= ust_kapasite * z[uretici]
    )
    
    # Üretici seçim kısıtı: Herhangi bir üründen üretim varsa, üretici seçilmiş olmalı
    for urun in gecerli_urun_ciftleri:
        solver.Add(x[(urun, uretici)] <= ust_kapasite * z[uretici])

# Amaç Fonksiyonu
objective = solver.Objective()
for (urun, uretici), var in x.items():
    fiyat = satis_fiyat[urun]
    olasilik = sales_probability[urun]
    maliyet = urun_uretici_dict.get((urun, uretici), 0)
    net_kar = (fiyat * olasilik) - maliyet
    objective.SetCoefficient(var, net_kar)

objective.SetMaximization()

status = solver.Solve()

def format_number(number):
    """Sayıları binlik ayracı olarak nokta kullanarak formatlar"""
    if isinstance(number, float):
        # Ondalık kısım varsa 2 basamak göster
        return f"{number:,.2f}".replace(",", ".")
    else:
        # Tam sayılar için ondalık gösterme
        return f"{number:,}".replace(",", ".")


if status == pywraplp.Solver.OPTIMAL:
    print("Optimal çözüm bulundu!\n")
    
    toplam_uretim = 0
    toplam_maliyet = 0
    toplam_beklenen_gelir = 0
    
    # Üretim kararlarını listeleme
    print("=== ÜRETİM KARARLARI ===")
    print("\nÜretilecek Ürünler:")
    uretilecek_urunler = [urun for urun in urunler if y[urun].solution_value() > 0]
    if uretilecek_urunler:
        for urun in uretilecek_urunler:
            print(f"- {urun}")
    else:
        print("Üretilecek ürün bulunmamaktadır.")
        
    print("\nÜretilmeyecek Ürünler:")
    uretilmeyecek_urunler = [urun for urun in urunler if y[urun].solution_value() == 0]
    if uretilmeyecek_urunler:
        for urun in uretilmeyecek_urunler:
            print(f"- {urun}")
    else:
        print("Üretilmeyecek ürün bulunmamaktadır.")
    
    print("\n=== ÜRETİCİ KULLANIM DURUMU ===")
    print("\nKullanılacak Üreticiler:")
    kullanilacak_ureticiler = [uretici for uretici in ureticiler if z[uretici].solution_value() > 0]
    if kullanilacak_ureticiler:
        for uretici in kullanilacak_ureticiler:
            toplam_uretim_miktari = sum(x[(urun, uretici)].solution_value() 
                                      for urun in urunler 
                                      if (urun, uretici) in x)
            print(f"- {uretici} (Toplam Üretim: {format_number(int(toplam_uretim_miktari))} adet)")
    else:
        print("Kullanılacak üretici bulunmamaktadır.")
        
    print("\nKullanılmayacak Üreticiler:")
    kullanilmayacak_ureticiler = [uretici for uretici in ureticiler if z[uretici].solution_value() == 0]
    if kullanilmayacak_ureticiler:
        for uretici in kullanilmayacak_ureticiler:
            print(f"- {uretici}")
    else:
        print("Kullanılmayacak üretici bulunmamaktadır.")
    
    print("\n=== DETAYLI ÜRETİM PLANI ===")
    # Ürünlerin detaylarını yazdırmak için
    for urun in uretilecek_urunler:
        urun_toplam_adet = 0
        urun_toplam_maliyet = 0
        urun_beklenen_gelir = 0
        
        satis_olasiligi = sales_probability.get(urun, 0)
        satis_fiyati = satis_fiyat.get(urun, 0)

        print(f"\nÜrün: {urun}")
        print(f"Satış Olasılığı: {format_number(satis_olasiligi)}")
        print(f"Satış Fiyatı: {format_number(satis_fiyati)}")
        
        # Ürün için tüm üreticilerdeki detayları hesapla
        for uretici in kullanilacak_ureticiler:
            if (urun, uretici) in x:
                uretim_adedi = x[(urun, uretici)].solution_value()
                if uretim_adedi > 0:
                    birim_maliyet = urun_uretici_dict[(urun, uretici)]
                    maliyet = uretim_adedi * birim_maliyet
                    beklenen_gelir = uretim_adedi * satis_fiyati * satis_olasiligi
                    
                    urun_toplam_adet += uretim_adedi
                    urun_toplam_maliyet += maliyet
                    urun_beklenen_gelir += beklenen_gelir
                    
                    print(f"  Üretici: {uretici}")
                    print(f"    Üretim Adedi: {format_number(int(uretim_adedi))}")
                    print(f"    Birim Maliyet: {format_number(birim_maliyet)}")
                    print(f"    Toplam Maliyet: {format_number(maliyet)}")
                    print(f"    Beklenen Gelir: {format_number(beklenen_gelir)}")
        
        if urun_toplam_adet > 0:
            urun_kar = urun_beklenen_gelir - urun_toplam_maliyet
            print(f"  Ürün Toplam Üretim: {format_number(int(urun_toplam_adet))}")
            print(f"  Ürün Toplam Maliyet: {format_number(urun_toplam_maliyet)}")
            print(f"  Ürün Beklenen Gelir: {format_number(urun_beklenen_gelir)}")
            print(f"  Ürün Beklenen Kâr: {format_number(urun_kar)}")
            
            toplam_uretim += urun_toplam_adet
            toplam_maliyet += urun_toplam_maliyet
            toplam_beklenen_gelir += urun_beklenen_gelir
    
    toplam_kar = toplam_beklenen_gelir - toplam_maliyet
    
    print("\n=== GENEL ÖZET ===")
    print(f"Toplam Üretim Adedi: {format_number(int(toplam_uretim))}")
    print(f"Toplam Maliyet: {format_number(toplam_maliyet)}")
    print(f"Toplam Beklenen Gelir: {format_number(toplam_beklenen_gelir)}")
    print(f"Toplam Beklenen Kâr: {format_number(toplam_kar)}")



    