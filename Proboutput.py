import pandas as pd
from ortools.linear_solver import pywraplp
pd.set_option('future.no_silent_downcasting', True)
import xlsxwriter

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

sales_probability = dict(
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
    olasilik = sales_probability[urun]
    maliyet = urun_uretici_dict.get((urun, uretici), 0)
    net_kar = (fiyat * olasilik) - maliyet
    
    # Çok küçük bir pozitif sayı ekleyerek sıfır üretimden kaçınmayı teşvik et
    objective.SetCoefficient(var, net_kar + 0.0001)

objective.SetMaximization()
    


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
    


# Katsayıları depolamak için boş listeler oluştur
data_rows = []

# Her ürün ve üretici kombinasyonu için katsayıları hesapla
for urun in urunler:
    fiyat = satis_fiyat[urun]
    olasilik = sales_probability[urun]
    
    for uretici in ureticiler:
        if (urun, uretici) in urun_uretici_dict:
            maliyet = urun_uretici_dict[(urun, uretici)]
            net_kar = (fiyat * olasilik) - maliyet + 0.0001
            
            # Her satır için veri sözlüğü oluştur
            row_data = {
                'Ürün': urun,
                'Üretici': uretici,
                'Satış Fiyatı': fiyat,
                'Satış Olasılığı': olasilik,
                'Birim Maliyet': maliyet,
                'Net Katsayı': net_kar,
                'Beklenen Gelir': fiyat * olasilik
            }
            data_rows.append(row_data)

# DataFrame oluştur
df_coefficients = pd.DataFrame(data_rows)

# Sütunları düzenle
df_coefficients = df_coefficients[[
    'Ürün', 
    'Üretici', 
    'Satış Fiyatı', 
    'Satış Olasılığı', 
    'Birim Maliyet', 
    'Beklenen Gelir',
    'Net Katsayı'
]]

# Sayısal değerleri formatla
format_dict = {
    'Satış Fiyatı': '{:.2f}',
    'Satış Olasılığı': '{:.2%}',
    'Birim Maliyet': '{:.2f}',
    'Beklenen Gelir': '{:.2f}',
    'Net Katsayı': '{:.2f}'
}

# DataFrame'i formatlı şekilde Excel'e kaydet
excel_file = 'urun_katsayilari.xlsx'
with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
    df_coefficients.to_excel(writer, sheet_name='Katsayılar', index=False)
    
    # Excel çalışma kitabı ve sayfasını al
    workbook = writer.book
    worksheet = writer.sheets['Katsayılar']
    
    # Formatlama için stil oluştur
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    percent_format = workbook.add_format({'num_format': '0.00%'})
    
    # Sütunlara format uygula
    worksheet.set_column('C:C', 12, number_format)  # Satış Fiyatı
    worksheet.set_column('D:D', 12, percent_format) # Satış Olasılığı
    worksheet.set_column('E:E', 12, number_format)  # Birim Maliyet
    worksheet.set_column('F:F', 12, number_format)  # Beklenen Gelir
    worksheet.set_column('G:G', 12, number_format)  # Net Katsayı
    
    # Sütun genişliklerini ayarla
    worksheet.set_column('A:A', 15)  # Ürün
    worksheet.set_column('B:B', 15)  # Üretici

print(f"Katsayılar '{excel_file}' dosyasına kaydedildi.")



