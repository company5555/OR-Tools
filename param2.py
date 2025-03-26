import pandas as pd
import numpy as np
import scipy.stats as stats

# CSV dosyasını noktalı virgülle okuma
urun_adet_data = pd.read_csv('Test.csv', sep=';')

# Her ürün için istatistiksel parametreleri hesaplama
urun_parametreleri = {}

for _, row in urun_adet_data.iterrows():
    urun = row['Ürünler']
    satis_adetleri = row[['2020', '2021', '2022', '2023', '2024']].values
    
    # Temel istatistikler
    ortalama = np.mean(satis_adetleri)
    standart_sapma = np.std(satis_adetleri)

    
    
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

print(sonuc_df)