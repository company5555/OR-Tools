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

# Üretici ve ürün bilgilerini birleştirme
products = ['T-shirt', 'Şort', 'Kazak']
producer_data = {
    'T-shirt': df_tshirt,
    'Şort': df_short,
    'Kazak': df_kazak
}

# Satış olasılıklarını ve fiyatlarını ayıklama
sales_prices = {}
sales_probabilities = {}
for _, row in df_sales.iterrows():
    product = row['Ürün']
    if product in products:
        sales_prices[product] = row['Satış Fiyatı']
        sales_probabilities[product] = row['Ortalama']
    else:
        print(f"Uyarı: '{product}' adlı ürün beklenen ürünler arasında değil!")

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
for product, df in producer_data.items():
    for _, row in df.iterrows():
        producer = row[df.columns[0]]
        solver.Add(x[producer, product] <= row['Kapasite'])

# Minimum üretim kısıtları
min_short_production = solver.Sum(x[producer, 'Şort'] for producer, product in x if product == 'Şort')
solver.Add(min_short_production >= 70000)

min_kazak_production = solver.Sum(x[producer, 'Kazak'] for producer, product in x if product == 'Kazak')
solver.Add(min_kazak_production >= 30000)

# Toplam üretim sınırı (örnek olarak 100000 adede kadar üretim)
total_production = solver.Sum(x[producer, product] for (producer, product) in x)
solver.Add(total_production <= 100000)

# Çözümü bulma
status = solver.Solve()

if status == pywraplp.Solver.OPTIMAL:
    print('Optimal çözüm bulundu:')
    for (producer, product), var in x.items():
        print(f'{producer} üreticisinin {product} üretimi: {var.solution_value()}')
    print(f'En yüksek kar: {solver.Objective().Value()}')
else:
    print('Optimal çözüm bulunamadı.')
