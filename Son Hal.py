import pandas as pd
from ortools.linear_solver import pywraplp
pd.set_option('future.no_silent_downcasting', True)

#Verileri Yükleme
file_path = "ORTEST.xlsx"

urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Satış")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
#"Toplam Maliyet" Satırını Hariç Tutma ve Ürün İsimlerini Normalize Etme

urun_kisit_data = urun_kisit_data[urun_kisit_data['Ürün'] != "Toplam Maliyet"]



#Ürünlerin tanımı
urunler = [urun for urun in urun_kisit_data['Ürün']]

ortalama_satis_olasiligi = dict(
    zip(
        urun_satis_data['Ürün'],
        urun_satis_data.iloc[:, 1:6].replace("-", 0).mean(axis=1)
    )
)

#Satış fiyatlarının tanımı
satis_fiyat = dict(
    zip(
        urun_satis_data['Ürün'],
        urun_satis_data['Satış Fiyatı']
    )
)

#Üreticilerin tanımlanması
ureticiler = [uretici for uretici in uretici_kapasite_data['Üretici']]

#Ürün ve Üretici Maliyet Bilgilerini Dictionary Olarak Hazırlama
urun_uretici_dict = {
    (row['Ürün'], row['Üretici']): row['Birim Maliyet']
    for _, row in urun_uretici_data.iterrows()
}

#Solver ve Değişkenlerin Tanımlanması
solver = pywraplp.Solver.CreateSolver('SCIP')

x = {}
for (urun, uretici) in urun_uretici_dict.keys():  # Sadece geçerli (Ürün, Üretici) çiftleri için değişken oluştur
    x[(urun, uretici)] = solver.IntVar(0, solver.infinity(), f'x_{urun}_{uretici}')




#Kısıtlar
#Ürün Alt ve Üst Üretim Sınırı
for urun in urunler:
    urun_alt = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
    urun_ust = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Üst Sınır'].values[0]
    
    gecerli_uretici_ciftleri = [uretici for uretici in ureticiler if (urun, uretici) in x]
    
    solver.Add(
        solver.Sum(x[(urun, uretici)] for uretici in gecerli_uretici_ciftleri) <= urun_ust
    )
# Önce alt sınırları düşük tutarak çözün
for urun in urunler:
    original_alt_sinir = urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'].values[0]
    urun_kisit_data.loc[urun_kisit_data['Ürün'] == urun, 'Üretim Alt Sınır'] = original_alt_sinir * 0.5  # %50 azalt


# Toplam Maliyet Kısıtı
toplam_maliyet_ust_sinir = float(toplam_maliyet_data[toplam_maliyet_data['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0])
print(f"Toplam Maliyet Üst Sınırı: {toplam_maliyet_ust_sinir}")

solver.Add(
    solver.Sum(
        x[(urun, uretici)] * urun_uretici_dict.get((urun, uretici), 0)
        for (urun, uretici) in x.keys()
    ) <= toplam_maliyet_ust_sinir
)

# Üretici Kapasite Kısıtı
for uretici in ureticiler:
    gecerli_urun_ciftleri = [urun for urun in urunler if (urun, uretici) in x]
    kapasite = uretici_kapasite_data.loc[uretici_kapasite_data['Üretici'] == uretici, 'Kapasite'].values[0]
    solver.Add(
        solver.Sum(x[(urun, uretici)] for urun in gecerli_urun_ciftleri) <= kapasite
    )

#Amaç Fonksiyonu
objective = solver.Objective()
for (urun, uretici), var in x.items():
    fiyat = satis_fiyat[urun]
    olasilik = ortalama_satis_olasiligi[urun]
    maliyet = urun_uretici_dict.get((urun, uretici), 0)
    net_kar = (fiyat * olasilik) - maliyet
    
    # Çok küçük bir pozitif sayı ekleyerek sıfır üretimden kaçınmayı teşvik et
    objective.SetCoefficient(var, net_kar + 0.0001)

objective.SetMaximization()
    


status = solver.Solve()

# Modelin Çözülmesi ve Sonuçların Çıktısı
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
    
    # Ürünlerin detaylarını yazdırmak için
    for urun in urunler:
        urun_toplam_adet = 0
        urun_toplam_maliyet = 0
        urun_beklenen_gelir = 0
        
        satis_olasiligi = ortalama_satis_olasiligi.get(urun, 0)
        satis_fiyati = satis_fiyat.get(urun, 0)

        print(f"\nÜrün: {urun}")
        print(f"Satış Olasılığı: {format_number(satis_olasiligi)}")
        print(f"Satış Fiyatı: {format_number(satis_fiyati)}")
        
        # Ürün için tüm üreticilerdeki detayları hesapla
        for uretici in ureticiler:
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

elif status == pywraplp.Solver.INFEASIBLE:
    print("Çözüm bulunamadı! Modelin kısıtları çelişkili olabilir.")
elif status == pywraplp.Solver.UNBOUNDED:
    print("Çözüm sınırsız! Modelde bir hata olabilir, kısıtlar yeterince sınırlayıcı değil.")
elif status == pywraplp.Solver.FEASIBLE:
    print("Feasable çözüm bulundu, ancak optimal çözüm değil.")
elif status == pywraplp.Solver.ABNORMAL:
    print("Çözümde anormal bir durum oluştu, çözüm algoritması hatalı olabilir.")
else:
    print("Çözüm bulunmadı! Çözüm algoritması çalıştırılmadı.")

