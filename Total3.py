import pandas as pd
from ortools.linear_solver import pywraplp

# Verileri Excel'den okuma
file_path = "ORTEST.xlsx"
df_urun_uretici = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
df_urun_kisit = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
df_urun_satis = pd.read_excel(file_path, sheet_name="Ürün - Satış")

# OR-Tools çözümleyicisini oluşturma
solver = pywraplp.Solver.CreateSolver("SCIP")
inf = solver.infinity()

# Ürün listesi
products = df_urun_uretici['Ürün'].unique()

# Satış fiyatlarını belirleme
sales_prices = {row['Ürün']: row['Satış Fiyatı'] for _, row in df_urun_satis.iterrows()}

# Karar değişkenleri: x[producer, product]
x = {}
for _, row in df_urun_uretici.iterrows():
    product = row['Ürün']
    producer = row['Üretici']
    capacity = row['Kapasite']
    x[producer, product] = solver.IntVar(0, capacity, f'x_{producer}_{product}')

# Amaç fonksiyonu: Kar maksimize etme
objective = solver.Objective()
for _, row in df_urun_uretici.iterrows():
    product = row['Ürün']
    producer = row['Üretici']
    unit_cost = row['Birim Maliyet']
    sales_price = sales_prices.get(product, 0)
    profit_per_unit = sales_price - unit_cost
    objective.SetCoefficient(x[producer, product], profit_per_unit)
objective.SetMaximization()

# Kapasite kısıtları
capacity_constraints = []
for _, row in df_urun_uretici.iterrows():
    product = row['Ürün']
    producer = row['Üretici']
    capacity = row['Kapasite']
    constraint = solver.Add(x[producer, product] <= capacity)
    capacity_constraints.append((constraint, f"{producer} kapasite kısıtı {product}"))

# Minimum ve maksimum üretim kısıtları
production_constraints = []
for product in products:
    product_constraints = df_urun_kisit[df_urun_kisit['Ürün'] == product]
    
    if not product_constraints.empty:
        lower_bound = product_constraints['Üretim Alt Sınır'].iloc[0]
        upper_bound = product_constraints['Üretim Üst Sınır'].iloc[0]
        
        total_production = solver.Sum(
            x[producer, product] 
            for producer in df_urun_uretici[df_urun_uretici['Ürün'] == product]['Üretici']
        )
        
        if pd.notna(lower_bound) and lower_bound > 0:
            constraint = solver.Add(total_production >= lower_bound)
            production_constraints.append((constraint, f"{product} Üretim Alt Sınırı"))
            
        if pd.notna(upper_bound) and upper_bound > 0:
            constraint = solver.Add(total_production <= upper_bound)
            production_constraints.append((constraint, f"{product} Üretim Üst Sınırı"))

# Toplam maliyet kısıtı
total_cost = solver.Sum(
    x[producer, product] * row['Birim Maliyet']
    for _, row in df_urun_uretici.iterrows()
    for producer in [row['Üretici']]
    for product in [row['Ürün']]
)

maliyet_siniri_row = df_urun_kisit[df_urun_kisit['Ürün'] == 'Toplam Maliyet']
if not maliyet_siniri_row.empty:
    maliyet_siniri = maliyet_siniri_row['Maliyet'].iloc[0]
    maliyet_constraint = solver.Add(total_cost <= maliyet_siniri)
else:
    print("Uyarı: Toplam maliyet sınırı bulunamadı!")
    maliyet_constraint = None

# Çözümü bulma
status = solver.Solve()

def print_production_possibilities():
    print("\nÜretim Olasılıkları:")
    print("-" * 50)
    for product in products:
        print(f"\n{product} için üretim detayları:")
        product_producers = df_urun_uretici[df_urun_uretici['Ürün'] == product]
        
        print(f"Satış Fiyatı: {sales_prices.get(product, 'Belirlenmemiş')}")
        
        for _, row in product_producers.iterrows():
            producer = row['Üretici']
            capacity = row['Kapasite']
            unit_cost = row['Birim Maliyet']
            profit = sales_prices.get(product, 0) - unit_cost
            
            print(f"\nÜretici: {producer}")
            print(f"Kapasite: {capacity}")
            print(f"Birim Maliyet: {unit_cost}")
            print(f"Birim Kar: {profit}")

if status == pywraplp.Solver.OPTIMAL:
    total_profit = solver.Objective().Value()
    
    if total_profit <= 0:
        print('Optimal çözüm bulunamadı: Kar negatif veya sıfır!')
        print(f'Hesaplanan kar: {total_profit}')
        print_production_possibilities()
    else:
        print('\nOptimal çözüm bulundu:')
        print('-' * 50)
        
        # Üretim miktarları
        product_totals = {}
        for (producer, product), var in x.items():
            production_amount = var.solution_value()
            if production_amount > 0:
                print(f'{producer} üreticisinin {product} üretimi: {production_amount}')
                product_totals[product] = product_totals.get(product, 0) + production_amount
        
        print('\nÜrün bazında toplam üretimler:')
        print('-' * 50)
        for product, total in product_totals.items():
            print(f'{product}: {total}')
        
        print(f'\nToplam kar: {total_profit}')
        print(f'Toplam maliyet: {total_cost.solution_value()}')
        
        print('\nÜretim olasılıkları ve kısıtlar:')
        print_production_possibilities()
        
else:
    print('\nOptimal çözüm bulunamadı!')
    print('\nAktif kısıtlar:')
    constraints_to_check = capacity_constraints + production_constraints
    if maliyet_constraint:
        constraints_to_check.append((maliyet_constraint, "Toplam maliyet sınırı"))
    
    for constraint, name in constraints_to_check:
        if constraint.Lb() > constraint.Ub():
            print(f"{name} etkin durumda: {constraint.Lb()} > {constraint.Ub()}")
    
    print('\nMevcut üretim olasılıkları:')
    print_production_possibilities()