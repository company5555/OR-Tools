import pandas as pd
from ortools.linear_solver import pywraplp
pd.set_option('future.no_silent_downcasting', True)

# 1. Verileri Yükleme
file_path = "ORTEST.xlsx"

urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Satış")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
toplam_maliyet_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
# 2. "Toplam Maliyet" Satırını Hariç Tutma ve Ürün İsimlerini Normalize Etme
urun_kisit_data = urun_kisit_data[urun_kisit_data['Ürün'] != "Toplam Maliyet"]

print(toplam_maliyet_data.index)



urunler = [urun for urun in urun_kisit_data['Ürün']]
print(urunler)
ortalama_satis_olasiligi = dict(
    zip(
        urun_satis_data['Ürün'],
        urun_satis_data.iloc[:, 1:6].replace("-", 0).mean(axis=1)
    )
)

print(ortalama_satis_olasiligi)
satis_fiyat = dict(
    zip(
        urun_satis_data['Ürün'],
        urun_satis_data['Satış Fiyatı']
    )
)
ureticiler = [uretici for uretici in uretici_kapasite_data['Üretici']]
print(ureticiler)
# 3. Ürün ve Üretici Maliyet Bilgilerini Dictionary Olarak Hazırlama
urun_uretici_dict = {
    (row['Ürün'], row['Üretici']): row['Birim Maliyet']
    for _, row in urun_uretici_data.iterrows()
}

# 4. Solver ve Değişkenlerin Tanımlanması
solver = pywraplp.Solver.CreateSolver('SCIP')

x = {}
for (urun, uretici) in urun_uretici_dict.keys():  # Sadece geçerli (Ürün, Üretici) çiftleri için değişken oluştur
    x[(urun, uretici)] = solver.IntVar(0, solver.infinity(), f'x_{urun}_{uretici}')




# 6. Kısıtlar
# Ürün Alt ve Üst Üretim Sınırı
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

# Çözüm başarılı olursa, kademeli olarak alt sınırları artırın
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

# 5. Amaç Fonksiyonu
# 5. Amaç Fonksiyonu
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
# Çözüm sonrası analiz
def analyze_solution():
    if status == pywraplp.Solver.OPTIMAL:
        print("\nÇözüm Analizi:")
        toplam_uretim = 0
        toplam_maliyet = 0
        beklenen_gelir = 0
        
        for (urun, uretici), var in x.items():
            miktar = var.solution_value()
            if miktar > 0:
                maliyet = urun_uretici_dict[(urun, uretici)]
                fiyat = satis_fiyat[urun]
                olasilik = ortalama_satis_olasiligi[urun]
                
                toplam_uretim += miktar
                toplam_maliyet += miktar * maliyet
                beklenen_gelir += miktar * fiyat * olasilik
                
                print(f"\nÜrün: {urun}, Üretici: {uretici}")
                print(f"  Üretim Miktarı: {miktar}")
                print(f"  Birim Maliyet: {maliyet}")
                print(f"  Toplam Maliyet: {miktar * maliyet}")
                print(f"  Beklenen Gelir: {miktar * fiyat * olasilik}")
        
        print(f"\nToplam Üretim: {toplam_uretim}")
        print(f"Toplam Maliyet: {toplam_maliyet}")
        print(f"Toplam Beklenen Gelir: {beklenen_gelir}")
        print(f"Beklenen Net Kâr: {beklenen_gelir - toplam_maliyet}")

# Çözümden sonra analizi yapın
analyze_solution()
# 7. Modelin Çözülmesi ve Sonuçların Çıktısı
status = solver.Solve()

if status == pywraplp.Solver.OPTIMAL:
    toplam_kar = 0
    print("Optimal çözüm bulundu!\n")
    
    # Ürünlerin detaylarını yazdırmak için
    for urun in urunler:
        toplam_adet = 0
        satis_olasiligi = ortalama_satis_olasiligi.get(urun, 0)
        satis_fiyati = satis_fiyat.get(urun, 0)

        # Ürün için tüm üreticilerdeki toplam üretim adedini hesapla
        for uretici in ureticiler:
            if (urun, uretici) in x:
                toplam_adet += x[(urun, uretici)].solution_value()

        # Çıktıyı yazdır
        print(f"Ürün: {urun}, Toplam Üretim Adeti: {int(toplam_adet)}, "
              f"Satış Olasılığı: {satis_olasiligi:.2f}, Satış Fiyatı: {satis_fiyati:.2f}")
    
    print(f"\nToplam Kar: {toplam_kar:.2f}")

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




