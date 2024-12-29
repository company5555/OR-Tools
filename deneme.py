import pandas as pd
from ortools.linear_solver import pywraplp

# Verileri Excel'den okuma
file_path = "ORTEST.xlsx"
df_urunler = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
df_sales = pd.read_excel(file_path, sheet_name="Ürün - Satış")
df_producers = pd.read_excel(file_path, sheet_name="Ürün - Üretici")

# OR-Tools çözümleyicisini oluşturma
solver = pywraplp.Solver.CreateSolver("SCIP")
inf = solver.infinity()

# Ürünler
products = df_sales['Ürün'].unique()

# Satış fiyatlarını belirleme
sales_prices = {row['Ürün']: row['Satış Fiyatı'] for _, row in df_sales.iterrows()}

# "-" sembollerini NaN olarak değiştirme
df_sales.replace("-", pd.NA, inplace=True)

# Satış olasılıklarını belirleme (sadece dolu olan yılların ortalamasını alarak)
sales_probabilities = {}
for product in products:
    product_data = df_sales[df_sales['Ürün'] == product]
    if not product_data.empty:
        # Sadece sayısal sütunları seçip NaN değerleri dikkate almadan ortalama alıyoruz
        sales_columns = product_data.iloc[:, 1:product_data.columns.get_loc("Satış Fiyatı")].apply(pd.to_numeric, errors='coerce')
        sales_probabilities[product] = sales_columns.mean(axis=1, skipna=True).values[0]
    else:
        sales_probabilities[product] = 0

# Kontrol amaçlı yazdırma
print("Satış Fiyatları:", sales_prices)
print("Satış Olasılıkları:", sales_probabilities)

# Karar değişkenleri (ikili indislerle): x[producer, product]
x = {}
for _, row in df_producers.iterrows():
    producer = row['Üretici']
    product = row['Ürün']
    x[producer, product] = solver.IntVar(0, row['Kapasite'], f'x_{producer}_{product}')

# Amaç fonksiyonu: Satış olasılığı * satış adeti * satış fiyatı - üretim maliyeti ile kar maksimize etme
objective = solver.Objective()
for _, row in df_producers.iterrows():
    producer = row['Üretici']
    product = row['Ürün']
    unit_cost = row['Birim Maliyet']
    unit_price = sales_prices.get(product, 0)
    sales_probability = sales_probabilities.get(product, 0)
    profit_per_unit = (unit_price * sales_probability) - unit_cost
    objective.SetCoefficient(x[producer, product], profit_per_unit)
objective.SetMaximization()

# Kısıtlar: Üretici kapasite kısıtları
capacity_constraints = []
for _, row in df_producers.iterrows():
    producer = row['Üretici']
    product = row['Ürün']
    constraint = solver.Add(x[producer, product] <= row['Kapasite'])
    capacity_constraints.append((constraint, f"{producer} kapasite kısıtı {product}"))

# Ürün bazlı maksimum üretim kısıtları
for product in df_urunler['Ürün']:
    max_production = df_urunler[df_urunler['Ürün'] == product]['Üretim Üst Sınır'].values[0]
    if max_production > 0:
        product_constraint = solver.Sum(x[producer, prod] for producer, prod in x if prod == product)
        solver.Add(product_constraint <= max_production)


# Minimum üretim kısıtları
for product in df_urunler['Ürün']:
    min_production = df_urunler[df_urunler['Ürün'] == product]['Üretim Alt Sınır'].values[0]
    if min_production > 0:
        product_constraint = solver.Sum(x[producer, product] for producer, prod in x if prod == product)
        solver.Add(product_constraint >= min_production)

# Toplam maliyet sınırı
total_cost = solver.Sum(x[producer, product] * row['Birim Maliyet'] for _, row in df_producers.iterrows() for producer, product in [(row['Üretici'], row['Ürün'])])
maliyet_siniri = df_urunler[df_urunler['Ürün'] == 'Toplam Maliyet']['Maliyet'].values[0]
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
