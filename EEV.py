import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp

# Parametreler
SIMULASYON_SAYISI = 1000  # RP için kullanılan senaryo sayısı
EEV_SIMULASYON_SAYISI = 1000  # EEV hesaplaması için test sayısı
np.random.seed(12)

# Excel'den veri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

urunler = [u for u in urun_kisit_data['Ürün'] if u != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# 1. RP (Stokastik) Çözüm ve Ortalama Kar
SIMULASYON_SONUCLARI = []
for i in range(SIMULASYON_SAYISI):
    np.random.seed(12 + i)
    sales_stochastic = {u: max(0, np.random.normal(p['ortalama'], p['std'])) for u, p in urun_param_dict.items()}
    urun_ust_kisit = {u: min(sales_stochastic[u], urun_kisit_data.loc[urun_kisit_data['Ürün'] == u, 'Üretim Üst Sınır'].values[0]) for u in urunler}

    solver = pywraplp.Solver.CreateSolver('SCIP')
    x = {(u, j): solver.IntVar(0, urun_ust_kisit[u], f'x_{u}_{j}') for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict}

    objective = solver.Objective()
    for (u, j), var in x.items():
        kar = satis_fiyat[u] - urun_uretici_dict[(u, j)]
        objective.SetCoefficient(var, kar)
    objective.SetMaximization()

    for u in urunler:
        solver.Add(sum(x[(u, j)] for j in ureticiler if (u, j) in x) <= urun_ust_kisit[u])

    for j in ureticiler:
        solver.Add(sum(x[(u, j)] for u in urunler if (u, j) in x) <= uretici_kapasite_dict.get(j, float('inf')))
        solver.Add(sum(x[(u, j)] for u in urunler if (u, j) in x) >= uretici_alt_kapasite_dict.get(j, 0))

    if solver.Solve() == pywraplp.Solver.OPTIMAL:
        x_values = {k: v.solution_value() for k, v in x.items()}
        gelir = sum(x_values[(u, j)] * satis_fiyat[u] for (u, j) in x)
        maliyet = sum(x_values[(u, j)] * urun_uretici_dict[(u, j)] for (u, j) in x)
        SIMULASYON_SONUCLARI.append(gelir - maliyet)
    else:
        SIMULASYON_SONUCLARI.append(0)

RP = np.mean(SIMULASYON_SONUCLARI)
print(f"RP (Stokastik çözüm ortalama karı): {RP:,.2f}")

# 2. EV çözümünü al (ortalama talep ile)
average_demand = {u: urun_param_dict[u]["ortalama"] for u in urunler}
urun_ust_kisit_EV = {u: min(average_demand[u], urun_kisit_data.loc[urun_kisit_data['Ürün'] == u, 'Üretim Üst Sınır'].values[0]) for u in urunler}

solver_ev = pywraplp.Solver.CreateSolver('SCIP')
x_ev = {(u, j): solver_ev.IntVar(0, urun_ust_kisit_EV[u], f'x_ev_{u}_{j}') for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict}

objective_ev = solver_ev.Objective()
for (u, j), var in x_ev.items():
    kar = satis_fiyat[u] - urun_uretici_dict[(u, j)]
    objective_ev.SetCoefficient(var, kar)
objective_ev.SetMaximization()

for u in urunler:
    solver_ev.Add(sum(x_ev[(u, j)] for j in ureticiler if (u, j) in x_ev) <= urun_ust_kisit_EV[u])

for j in ureticiler:
    solver_ev.Add(sum(x_ev[(u, j)] for u in urunler if (u, j) in x_ev) <= uretici_kapasite_dict.get(j, float('inf')))
    solver_ev.Add(sum(x_ev[(u, j)] for u in urunler if (u, j) in x_ev) >= uretici_alt_kapasite_dict.get(j, 0))

solver_ev.Solve()
ev_plan = {(u, j): x_ev[(u, j)].solution_value() for (u, j) in x_ev}

# 3. EV çözümünü senaryo bazlı test et (EEV)
eev_karlar = []
for i in range(EEV_SIMULASYON_SAYISI):
    np.random.seed(500 + i)
    talep_senaryosu = {u: max(0, np.random.normal(urun_param_dict[u]["ortalama"], urun_param_dict[u]["std"])) for u in urunler}
    toplam_satis = {u: min(sum(ev_plan.get((u, j), 0) for j in ureticiler), talep_senaryosu[u]) for u in urunler}
    gelir = sum(toplam_satis[u] * satis_fiyat[u] for u in urunler)
    maliyet = sum(ev_plan[(u, j)] * urun_uretici_dict[(u, j)] for (u, j) in ev_plan)
    eev_karlar.append(gelir - maliyet)

EEV = np.mean(eev_karlar)
VSS = EEV - RP
VSS_orani = (VSS / EEV) * 100

print(f"EEV (Ortalama talebe göre alınan kararların senaryo performansı): {EEV:,.2f}")
print(f"VSS (EEV - RP): {VSS:,.2f} ₺")
print(f"VSS Oranı: %{VSS_orani:.2f}")
