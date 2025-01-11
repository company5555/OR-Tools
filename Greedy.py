import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

# Verileri Yükleme
def load_data(file_path):
    urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
    urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Satış")
    urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
    uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
    
    # "Toplam Maliyet" satırını hariç tutma
    toplam_maliyet = float(urun_kisit_data[urun_kisit_data['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0])
    urun_kisit_data = urun_kisit_data[urun_kisit_data['Ürün'] != "Toplam Maliyet"]
    
    return urun_kisit_data, urun_satis_data, urun_uretici_data, uretici_kapasite_data, toplam_maliyet

def calculate_coefficients(urun_satis_data, urun_uretici_data):
    # Satış olasılıklarını hesapla
    sales_probability = dict(
        zip(
            urun_satis_data['Ürün'],
            urun_satis_data.iloc[:, 1:6].replace("-", 0).mean(axis=1)
        )
    )
    
    # Satış fiyatlarını al
    satis_fiyat = dict(
        zip(
            urun_satis_data['Ürün'],
            urun_satis_data['Satış Fiyatı']
        )
    )
    
    # Katsayıları hesapla (kar potansiyeli)
    coefficients = {}
    for _, row in urun_uretici_data.iterrows():
        urun = row['Ürün']
        uretici = row['Üretici']
        maliyet = row['Birim Maliyet']
        net_kar = (satis_fiyat[urun] * sales_probability[urun]) - maliyet
        coefficients[(urun, uretici)] = {
            'net_kar': net_kar,
            'birim_maliyet': maliyet
        }
    
    return coefficients, sales_probability, satis_fiyat

def greedy_optimization(file_path):
    # Verileri yükle
    urun_kisit_data, urun_satis_data, urun_uretici_data, uretici_kapasite_data, toplam_maliyet = load_data(file_path)
    
    # Katsayıları hesapla
    coefficients, sales_probability, satis_fiyat = calculate_coefficients(urun_satis_data, urun_uretici_data)
    
    # Üretici kapasitelerini dictionary'e al
    uretici_kapasiteleri = {}
    for _, row in uretici_kapasite_data.iterrows():
        uretici_kapasiteleri[row['Üretici']] = {
            'kalan_kapasite': row['Üst Kapasite'],
            'alt_kapasite': row['Alt Kapasite'],
            'ust_kapasite': row['Üst Kapasite']
        }
    
    # Ürün üretim sınırlarını dictionary'e al
    urun_sinirlari = {}
    for _, row in urun_kisit_data.iterrows():
        urun_sinirlari[row['Ürün']] = {
            'kalan_uretim': row['Üretim Üst Sınır'],
            'alt_sinir': row['Üretim Alt Sınır'],
            'ust_sinir': row['Üretim Üst Sınır']
        }
    
    # Sonuçları tutacak dictionary
    uretim_plani = {}
    kalan_toplam_maliyet = toplam_maliyet
    
    # Katsayıları sırala
    sorted_coefficients = sorted(
        coefficients.items(),
        key=lambda x: x[1]['net_kar'],
        reverse=True
    )
    
    # Greedy algoritma
    for (urun, uretici), coefficient in sorted_coefficients:
        if kalan_toplam_maliyet <= 0:
            break
            
        max_uretim = min(
            urun_sinirlari[urun]['kalan_uretim'],
            uretici_kapasiteleri[uretici]['kalan_kapasite'],
            int(kalan_toplam_maliyet / coefficient['birim_maliyet'])
        )
        
        if max_uretim > 0:
            uretim_plani[(urun, uretici)] = max_uretim
            
            # Kapasiteleri güncelle
            urun_sinirlari[urun]['kalan_uretim'] -= max_uretim
            uretici_kapasiteleri[uretici]['kalan_kapasite'] -= max_uretim
            kalan_toplam_maliyet -= max_uretim * coefficient['birim_maliyet']
    
    # Sonuçları yazdır
    print_results(uretim_plani, coefficients, sales_probability, satis_fiyat)

def format_number(number):
    """Sayıları binlik ayracı olarak nokta kullanarak formatlar"""
    if isinstance(number, float):
        return f"{number:,.2f}".replace(",", ".")
    else:
        return f"{number:,}".replace(",", ".")

def print_results(uretim_plani, coefficients, sales_probability, satis_fiyat):
    toplam_uretim = 0
    toplam_maliyet = 0
    toplam_beklenen_gelir = 0
    
    print("=== DETAYLI ÜRETİM PLANI ===")
    
    # Ürünlere göre grupla
    urun_bazli_plan = {}
    for (urun, uretici), miktar in uretim_plani.items():
        if urun not in urun_bazli_plan:
            urun_bazli_plan[urun] = []
        urun_bazli_plan[urun].append((uretici, miktar))
    
    # Her ürün için detayları yazdır
    for urun, uretimler in urun_bazli_plan.items():
        urun_toplam_adet = 0
        urun_toplam_maliyet = 0
        urun_beklenen_gelir = 0
        
        print(f"\nÜrün: {urun}")
        print(f"Satış Olasılığı: {format_number(sales_probability[urun])}")
        print(f"Satış Fiyatı: {format_number(satis_fiyat[urun])}")
        
        for uretici, miktar in uretimler:
            birim_maliyet = coefficients[(urun, uretici)]['birim_maliyet']
            maliyet = miktar * birim_maliyet
            beklenen_gelir = miktar * satis_fiyat[urun] * sales_probability[urun]
            
            print(f"  Üretici: {uretici}")
            print(f"    Üretim Adedi: {format_number(int(miktar))}")
            print(f"    Birim Maliyet: {format_number(birim_maliyet)}")
            print(f"    Toplam Maliyet: {format_number(maliyet)}")
            print(f"    Beklenen Gelir: {format_number(beklenen_gelir)}")
            
            urun_toplam_adet += miktar
            urun_toplam_maliyet += maliyet
            urun_beklenen_gelir += beklenen_gelir
        
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

# Çalıştırma
file_path = "ORTEST.xlsx"
greedy_optimization(file_path)