import pandas as pd
from ortools.linear_solver import pywraplp

# Verileri Excel'den okuma
file_path = "ORTEST.xlsx"
df_urunler = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
df_sales = pd.read_excel(file_path, sheet_name="Ürün - Satış")
df_producers = pd.read_excel(file_path, sheet_name="Ürün - Üretici")

# Debug için verileri yazdırma
print("\nÜrün üretim sınırları:")
print(df_urunler[['Ürün', 'Üretim Alt Sınır', 'Üretim Üst Sınır']])

# Satış fiyatlarını belirleme
sales_prices = {row['Ürün']: row['Satış Fiyatı'] for _, row in df_sales.iterrows()}

# OR-Tools çözümleyicisini oluşturma
solver = pywraplp.Solver.CreateSolver("SCIP")
if not solver:
    print("Solver oluşturulamadı!")
    exit(1)

# Ürünler
products = df_sales['Ürün'].unique()

# Satış olasılıklarını belirleme
df_sales.replace("-", pd.NA, inplace=True)
sales_probabilities = {}
for product in products:
    product_data = df_sales[df_sales['Ürün'] == product]
    if not product_data.empty:
        sales_columns = product_data.select_dtypes(include=['float64', 'int64']).iloc[:, :-1]
        non_zero_mean = sales_columns.replace(0, pd.NA).mean(axis=1, skipna=True)
        sales_probabilities[product] = non_zero_mean.iloc[0] if not non_zero_mean.empty else 0
    else:
        sales_probabilities[product] = 0

# Karar değişkenleri
x = {}
for _, row in df_producers.iterrows():
    producer = row['Üretici']
    product = row['Ürün']
    if pd.notna(row['Kapasite']):
        x[producer, product] = solver.IntVar(0, int(row['Kapasite']), f'x_{producer}_{product}')

# Amaç fonksiyonu
objective = solver.Objective()
for _, row in df_producers.iterrows():
    producer = row['Üretici']
    product = row['Ürün']
    if (producer, product) in x:
        unit_cost = float(row['Birim Maliyet'])
        unit_price = float(sales_prices.get(product, 0))
        sales_probability = float(sales_probabilities.get(product, 0))
        profit_per_unit = (unit_price * sales_probability) - unit_cost
        objective.SetCoefficient(x[producer, product], profit_per_unit)
objective.SetMaximization()

# Üretici kapasite kısıtları
for _, row in df_producers.iterrows():
    producer = row['Üretici']
    product = row['Ürün']
    if (producer, product) in x:
        solver.Add(x[producer, product] <= row['Kapasite'])

# Ürün bazlı kısıtlar - ÜST SINIR DÜZELTMESİ
print("\nUygulanan ürün kısıtları:")
for _, row in df_urunler.iterrows():
    product = row['Ürün']
    if product != 'Toplam Maliyet':
        product_vars = [x[producer, prod] for (producer, prod) in x.keys() if prod == product]
        
        if product_vars:
            total_production = solver.Sum(product_vars)
            
            # Üst sınır kontrolü ve kısıtı
            if pd.notna(row['Üretim Üst Sınır']):
                upper_bound = float(row['Üretim Üst Sınır'])
                if upper_bound > 0:
                    constraint = solver.Add(total_production <= upper_bound)
                    print(f"{product} için üst sınır kısıtı: <= {upper_bound}")
            
            # Alt sınır kontrolü ve kısıtı
            if pd.notna(row['Üretim Alt Sınır']):
                lower_bound = float(row['Üretim Alt Sınır'])
                if lower_bound > 0:
                    constraint = solver.Add(total_production >= lower_bound)
                    print(f"{product} için alt sınır kısıtı: >= {lower_bound}")

# Toplam maliyet kısıtı
total_cost = solver.Sum(x[producer, product] * row['Birim Maliyet'] 
                       for _, row in df_producers.iterrows() 
                       for producer, product in [(row['Üretici'], row['Ürün'])] 
                       if (producer, product) in x)

maliyet_siniri = float(df_urunler[df_urunler['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0])
solver.Add(total_cost <= maliyet_siniri)

# Çözümü bulma
status = solver.Solve()

# Sonuçları yazdırma
if status == pywraplp.Solver.OPTIMAL:
    print('\nOPTİMAL ÇÖZÜM DETAYLARI:')
    print('=' * 80)
    
    # Ürün bazlı sonuçlar için sözlükler
    urun_toplam_uretim = {}
    urun_toplam_maliyet = {}
    urun_toplam_gelir = {}
    
    # Her üretici ve ürün için sonuçları hesapla
    for (producer, product), var in x.items():
        production = var.solution_value()
        if production > 0:
            producer_row = df_producers[
                (df_producers['Üretici'] == producer) & 
                (df_producers['Ürün'] == product)
            ].iloc[0]
            
            if product not in urun_toplam_uretim:
                urun_toplam_uretim[product] = 0
                urun_toplam_maliyet[product] = 0
                urun_toplam_gelir[product] = 0
            
            uretim = production
            birim_maliyet = producer_row['Birim Maliyet']
            maliyet = uretim * birim_maliyet
            if product in sales_prices:
                satis_fiyati = sales_prices[product]
                satis_olasiligi = sales_probabilities[product]
                beklenen_gelir = uretim * satis_fiyati * satis_olasiligi
            else:
                print(f"UYARI: {product} için satış fiyatı bulunamadı!")
                continue
            
            urun_toplam_uretim[product] += uretim
            urun_toplam_maliyet[product] += maliyet
            urun_toplam_gelir[product] += beklenen_gelir
            
            print(f"\nÜretici: {producer} - Ürün: {product}")
            print(f"  Üretim Miktarı: {uretim:,.0f}")
            print(f"  Birim Maliyet: {birim_maliyet:,.2f}")
            print(f"  Toplam Maliyet: {maliyet:,.2f}")
            print(f"  Satış Fiyatı: {satis_fiyati:,.2f}")
            print(f"  Satış Olasılığı: {satis_olasiligi:,.2%}")
            print(f"  Beklenen Gelir: {beklenen_gelir:,.2f}")
    
    # Ürün bazlı özet
    print('\nÜRÜN BAZLI ÖZET:')
    print('=' * 80)
    genel_toplam_uretim = 0
    genel_toplam_maliyet = 0
    genel_toplam_gelir = 0
    
    for product in urun_toplam_uretim.keys():
        uretim = urun_toplam_uretim[product]
        maliyet = urun_toplam_maliyet[product]
        gelir = urun_toplam_gelir[product]
        kar = gelir - maliyet
        
        # Üst sınır kontrolü
        urun_ust_sinir = df_urunler[df_urunler['Ürün'] == product]['Üretim Üst Sınır'].values[0]
        
        print(f"\nÜrün: {product}")
        print(f"  Toplam Üretim: {uretim:,.0f} (Üst Sınır: {urun_ust_sinir:,.0f})")
        print(f"  Toplam Maliyet: {maliyet:,.2f}")
        print(f"  Beklenen Toplam Gelir: {gelir:,.2f}")
        print(f"  Beklenen Kar: {kar:,.2f}")
        
        genel_toplam_uretim += uretim
        genel_toplam_maliyet += maliyet
        genel_toplam_gelir += gelir
    
    print('\nGENEL ÖZET:')
    print('=' * 80)
    print(f"Toplam Üretim: {genel_toplam_uretim:,.0f}")
    print(f"Toplam Maliyet: {genel_toplam_maliyet:,.2f}")
    print(f"Beklenen Toplam Gelir: {genel_toplam_gelir:,.2f}")
    print(f"Beklenen Toplam Kar: {genel_toplam_gelir - genel_toplam_maliyet:,.2f}")
    
elif status == pywraplp.Solver.FEASIBLE:
    print('Uygun bir çözüm bulundu, ancak optimal olmayabilir.')
elif status == pywraplp.Solver.INFEASIBLE:
    print('Problem çözülemez (kısıtlar çelişiyor olabilir).')
else:
    print(f'Optimal çözüm bulunamadı. Solver durumu: {status}')