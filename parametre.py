import pandas as pd
import numpy as np
import scipy.stats as stats

# Veri yükleme
file_path = "ORTEST.xlsx"

# İlk satırı başlık olarak kullanmadan veriyi okuma
urun_adet_data = pd.read_excel(file_path, sheet_name="Ürün - Adet", header=None)

# İlk satırı sütun isimleri olarak ayarlama
urun_adet_data.columns = urun_adet_data.iloc[0]
urun_adet_data = urun_adet_data.drop(urun_adet_data.index[0]).reset_index(drop=True)

# Sütun isimlerini ve veri türünü düzenleme
urun_adet_data.columns = ['Ürünler'] + list(urun_adet_data.columns[1:])

# NaN değerlerini ve float'ları işleme
for col in ['2020', '2021', '2022', '2023', '2024']:
    urun_adet_data[col] = pd.to_numeric(urun_adet_data[col], errors='coerce').fillna(0).astype(int)

# Her ürün için istatistiksel parametreleri hesaplama
urun_parametreleri = {}

for _, row in urun_adet_data.iterrows():
    urun = row['Ürünler']
    satis_adetleri = row[['2020', '2021', '2022', '2023', '2024']].values
    
    # Temel istatistikler
    ortalama = np.mean(satis_adetleri)
    standart_sapma = np.std(satis_adetleri)
    
    # Normallik testi
    _, normallik_p_degeri = stats.normaltest(satis_adetleri)
    
    # Güven aralığı hesaplama (95% güven düzeyi)
    guven_araligi = stats.t.interval(alpha=0.95, 
                                     df=len(satis_adetleri)-1, 
                                     loc=ortalama, 
                                     scale=stats.sem(satis_adetleri))
    
    urun_parametreleri[urun] = {
        'Ortalama Satış Adedi': ortalama,
        'Standart Sapma': standart_sapma,
        'Varyans': standart_sapma**2,
        'Normallik P-Değeri': normallik_p_degeri,
        'Güven Aralığı Alt Sınır': guven_araligi[0],
        'Güven Aralığı Üst Sınır': guven_araligi[1],
        'Satış Adetleri': list(satis_adetleri)
    }

# Sonuçları ekrana yazdırma
print("Ürün Satış Parametreleri:\n")
for urun, parametreler in urun_parametreleri.items():
    print(f"Ürün: {urun}")
    for parametre, deger in parametreler.items():
        if parametre != 'Satış Adetleri':
            print(f"  {parametre}: {deger:.2f}")
    print(f"  Satış Adetleri: {parametreler['Satış Adetleri']}\n")

# Sonuçları DataFrame olarak kaydetme
sonuc_df = pd.DataFrame.from_dict(urun_parametreleri, orient='index')
sonuc_df.to_excel('urun_satis_parametreleri.xlsx')