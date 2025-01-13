import pandas as pd
pd.set_option('future.no_silent_downcasting', True)

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

def calculate_production_plan(urun_kisit_data, uretici_kapasite_data, coefficients, toplam_maliyet):
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
    
    return uretim_plani

def create_excel_report(uretim_plani, coefficients, sales_probability, satis_fiyat):
    # Dictionary to store results for each product
    results = {
        'Product': [],
        'Total Production Amount': [],
        'Total Income': [],
        'Total Cost': [],
        'Profit': []
    }
    
    # Group by product
    urun_bazli_plan = {}
    for (urun, uretici), miktar in uretim_plani.items():
        if urun not in urun_bazli_plan:
            urun_bazli_plan[urun] = []
        urun_bazli_plan[urun].append((uretici, miktar))
    
    # Calculate totals for each product
    for urun, uretimler in urun_bazli_plan.items():
        urun_toplam_adet = 0
        urun_toplam_maliyet = 0
        urun_beklenen_gelir = 0
        
        for uretici, miktar in uretimler:
            birim_maliyet = coefficients[(urun, uretici)]['birim_maliyet']
            maliyet = miktar * birim_maliyet
            beklenen_gelir = miktar * satis_fiyat[urun] * sales_probability[urun]
            
            urun_toplam_adet += miktar
            urun_toplam_maliyet += maliyet
            urun_beklenen_gelir += beklenen_gelir
        
        urun_kar = urun_beklenen_gelir - urun_toplam_maliyet
        
        # Add to results dictionary
        results['Product'].append(urun)
        results['Total Production Amount'].append(int(urun_toplam_adet))
        results['Total Income'].append(urun_beklenen_gelir)
        results['Total Cost'].append(urun_toplam_maliyet)
        results['Profit'].append(urun_kar)
    
    # Create DataFrame
    df = pd.DataFrame(results)
    
    # Add total row
    totals = pd.Series({
        'Product': 'TOTAL',
        'Total Production Amount': df['Total Production Amount'].sum(),
        'Total Income': df['Total Income'].sum(),
        'Total Cost': df['Total Cost'].sum(),
        'Profit': df['Profit'].sum()
    })
    
    df = pd.concat([df, pd.DataFrame([totals])], ignore_index=True)
    
    # Export to Excel
    df.to_excel('production_results.xlsx', index=False, float_format='%.2f')
    return df

def main(file_path):
    # Load data
    urun_kisit_data, urun_satis_data, urun_uretici_data, uretici_kapasite_data, toplam_maliyet = load_data(file_path)
    
    # Calculate coefficients
    coefficients, sales_probability, satis_fiyat = calculate_coefficients(urun_satis_data, urun_uretici_data)
    
    # Calculate production plan
    uretim_plani = calculate_production_plan(urun_kisit_data, uretici_kapasite_data, coefficients, toplam_maliyet)
    
    # Create and export Excel report
    df = create_excel_report(uretim_plani, coefficients, sales_probability, satis_fiyat)
    print("Results have been exported to 'production_results.xlsx'")
    print("\nSummary of results:")
    print(df.to_string(index=False))

if __name__ == "__main__":
    file_path = "ORTEST.xlsx"
    main(file_path)