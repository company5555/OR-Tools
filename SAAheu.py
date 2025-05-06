import pandas as pd
import numpy as np
import time
from ortools.linear_solver import pywraplp
from tqdm import tqdm

# ==========================================
# 1. YARDIMCI FONKSİYONLAR
# ==========================================

def generate_random_scenarios(urun_param_dict, sim_sayisi, seed=None):
    if seed is not None:
        np.random.seed(seed)
    return {
        urun: [max(0, int(round(np.random.normal(p["ortalama"], p["std"]))))
               for _ in range(sim_sayisi)]
        for urun, p in urun_param_dict.items()
    }

def solve_production_model(sales_scenarios, urunler, ureticiler, satis_fiyat,
                           urun_uretici_dict, uretici_kapasite_dict,
                           uretici_alt_kapasite_dict, sim_sayisi):

    solver = pywraplp.Solver.CreateSolver('SCIP')
    if solver is None:
        raise RuntimeError("Solver oluşturulamadı.")

    x = {
        (urun, uretici): solver.IntVar(0, sum(sales_scenarios[urun]) // sim_sayisi + 100,
                                       f'x_{urun}_{uretici}')
        for urun in urunler for uretici in ureticiler if (urun, uretici) in urun_uretici_dict
    }

    satilan = {
        (urun, s): solver.NumVar(0, solver.infinity(), f'satilan_{urun}_{s}')
        for urun in urunler for s in range(sim_sayisi)
    }

    for urun in urunler:
        toplam_uretim = sum(x[(urun, uretici)] for uretici in ureticiler if (urun, uretici) in x)
        for s in range(sim_sayisi):
            solver.Add(satilan[(urun, s)] <= toplam_uretim)
            solver.Add(satilan[(urun, s)] <= sales_scenarios[urun][s])

    for uretici in ureticiler:
        toplam = sum(x[(urun, uretici)] for urun in urunler if (urun, uretici) in x)
        solver.Add(toplam <= uretici_kapasite_dict.get(uretici, float('inf')))
        solver.Add(toplam >= uretici_alt_kapasite_dict.get(uretici, 0))

    objective = solver.Objective()
    for urun in urunler:
        for s in range(sim_sayisi):
            objective.SetCoefficient(satilan[(urun, s)], satis_fiyat[urun] / sim_sayisi)
    for (urun, uretici), var in x.items():
        objective.SetCoefficient(var, -urun_uretici_dict[(urun, uretici)])
    objective.SetMaximization()

    status = solver.Solve()
    if status == pywraplp.Solver.OPTIMAL:
        plan = {(urun, uretici): x[(urun, uretici)].solution_value()
                for (urun, uretici) in x}
        return plan
    else:
        return None

def evaluate_plan(plan, test_scenarios, urunler, ureticiler, satis_fiyat, urun_uretici_dict, sim_sayisi):
    total_profit = 0
    for urun in urunler:
        uretim_miktari = sum(plan.get((urun, uretici), 0) for uretici in ureticiler)
        for s in range(sim_sayisi):
            satis = min(uretim_miktari, test_scenarios[urun][s])
            total_profit += (satis * satis_fiyat[urun] / sim_sayisi)
    total_cost = sum(plan[(urun, uretici)] * urun_uretici_dict[(urun, uretici)]
                     for (urun, uretici) in plan)
    return total_profit - total_cost

# ==========================================
# 2. ANA AKIŞ – SAA YAKLAŞIMI
# ==========================================

# Dosya ve veri okuma
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

urunler = [urun for urun in urun_kisit_data['Ürün'] if urun != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))

satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# SAA Parametreleri
NUM_GROUPS = 5
NUM_EVALUATION = 100
SIMULASYON_SAYISI = 1000

best_plan = None
best_profit = -float('inf')

print("SAA başlatılıyor...\n")
for g in tqdm(range(NUM_GROUPS), desc="SAA Grupları"):
    scenarios = generate_random_scenarios(urun_param_dict, SIMULASYON_SAYISI, seed=g)
    plan = solve_production_model(
        scenarios, urunler, ureticiler, satis_fiyat, urun_uretici_dict,
        uretici_kapasite_dict, uretici_alt_kapasite_dict, SIMULASYON_SAYISI)

    if plan is None:
        continue

    evaluation_profits = []
    for e in range(NUM_EVALUATION):
        test_scenarios = generate_random_scenarios(urun_param_dict, SIMULASYON_SAYISI, seed=100+e)
        profit = evaluate_plan(plan, test_scenarios, urunler, ureticiler, satis_fiyat, urun_uretici_dict, SIMULASYON_SAYISI)
        evaluation_profits.append(profit)

    avg_profit = np.mean(evaluation_profits)
    if avg_profit > best_profit:
        best_profit = avg_profit
        best_plan = plan

# ==========================================
# 3. SONUÇLAR
# ==========================================

print("\n=== En İyi Plan ===")
for urun in urunler:
    toplam_uretim = sum(best_plan.get((urun, uretici), 0) for uretici in ureticiler)
    print(f"{urun}: {toplam_uretim}")

print(f"\nEn iyi planın ortalama karı (SAA test grupları üzerinde): {best_profit:,.2f}")
