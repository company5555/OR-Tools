import pandas as pd
from ortools.linear_solver import pywraplp

# Verileri Excel'den okuma
file_path = "ORTEST.xlsx"
df_urunler = pd.read_excel(file_path, sheet_name="Ürün Kısıt")
df_tshirt = pd.read_excel(file_path, sheet_name="T-shirt Üretici")
df_short = pd.read_excel(file_path, sheet_name="Şort Üretici")
df_kazak = pd.read_excel(file_path, sheet_name="Kazak Üretici")
df_sales = pd.read_excel(file_path, sheet_name="Ürün Satış")

# OR-Tools çözümleyicisini oluşturma
solver = pywraplp.Solver.CreateSolver("SCIP")
inf = solver.infinity()

# Ürünler
products = ['T-shirt', 'Şort', 'Kazak']

# Satış fiyatlarını belirleme
sales_prices = {row['Ürün']: row['Satış Fiyatı'] for _, row in df_sales.iterrows() if row['Ürün'] in products}

# Satış olasılıklarını belirleme (ortalama alarak)
sales_probabilities = {}
for product in products:
    product_data = df_sales[df_sales['Ürün'] == product]
    if not product_data.empty:
        sales_columns = product_data.iloc[:, 1:product_data.columns.get_loc("Satış Fiyatı")]
        sales_probabilities[product] = sales_columns.mean(axis=1).values[0]
    else:
        sales_probabilities[product] = 0

# Kontrol amaçlı yazdırma
print("Satış Fiyatları:", sales_prices)
print("Satış Olasılıkları:", sales_probabilities)

# Üretici ve ürün bilgilerini birleştirme
producer_data = {
    'T-shirt': df_tshirt,
    'Şort': df_short,
    'Kazak': df_kazak
}

# Karar değişkenleri (ikili indislerle): x[producer, product]
x = {}
for product, df in producer_data.items():
    for _, row in df.iterrows():
        producer = row[df.columns[0]]
        x[producer, product] = solver.IntVar(0, row['Kapasite'], f'x_{producer}_{product}')

# Amaç fonksiyonu: Satış olasılığı * satış adeti * satış fiyatı - üretim maliyeti ile kar maksimize etme
objective = solver.Objective()
for product, df in producer_data.items():
    for _, row in df.iterrows():
        producer = row[df.columns[0]]
        unit_cost = row['Birim Maliyet']
        unit_price = sales_prices.get(product, 0)
        sales_probability = sales_probabilities.get(product, 0)
        profit_per_unit = (unit_price * sales_probability) - unit_cost
        objective.SetCoefficient(x[producer, product], profit_per_unit)
objective.SetMaximization()

# Kısıtlar: Üretici kapasite kısıtları
capacity_constraints = []
for product, df in producer_data.items():
    for _, row in df.iterrows():
        producer = row[df.columns[0]]
        constraint = solver.Add(x[producer, product] <= row['Kapasite'])
        capacity_constraints.append((constraint, f"{producer} kapasite kısıtı {product}"))

# Minimum üretim kısıtları
min_short_production = solver.Sum(x[producer, 'Şort'] for producer, product in x if product == 'Şort')
short_constraint = solver.Add(min_short_production >= 70000)

min_kazak_production = solver.Sum(x[producer, 'Kazak'] for producer, product in x if product == 'Kazak')
kazak_constraint = solver.Add(min_kazak_production >= 30000)

# Toplam maliyet sınırı
total_cost = solver.Sum(x[producer, product] * row['Birim Maliyet'] for product, df in producer_data.items() for _, row in df.iterrows() for producer in [row[df.columns[0]]])
maliyet_siniri = df_urunler[df_urunler['Ürün'] == 'Toplam Maliyet']['Maliyet Sınırı'].values[0]
maliyet_constraint = solver.Add(total_cost <= maliyet_siniri)

# Çözümü bulma
status = solver.Solve()

if status == pywraplp.Solver.OPTIMAL:
    print('Optimal çözüm bulundu:')
    for (producer, product), var in x.items():
        print(f'{producer} üreticisinin {product} üretimi: {var.solution_value()}')
    print(f'En yüksek kar: {solver.Objective().Value()}')
    print("Satış Olasılıkları:", sales_probabilities)
else:
    print('Optimal çözüm bulunamadı.')
    # Aktif kısıtları kontrol etme
    print("Aktif kısıtlar:")
    for constraint, name in capacity_constraints:
        if constraint.Lb() > constraint.Ub():
            print(f"{name} etkin durumda: {constraint.Lb()} > {constraint.Ub()}")
    if short_constraint.Lb() > short_constraint.Ub():
        print(f"Şort minimum üretim kısıtı etkin: {short_constraint.Lb()} > {short_constraint.Ub()}")
    if kazak_constraint.Lb() > kazak_constraint.Ub():
        print(f"Kazak minimum üretim kısıtı etkin: {kazak_constraint.Lb()} > {kazak_constraint.Ub()}")
    if maliyet_constraint.Lb() > maliyet_constraint.Ub():
        print(f"Toplam maliyet sınırı etkin: {maliyet_constraint.Lb()} > {maliyet_constraint.Ub()}")
