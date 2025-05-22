import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
from tqdm import tqdm
import time

# === Parametreler ===
SIMULASYON_SAYISI = 10  # Senaryo sayısı (Sample Average Approximation)
# === Excel'den veri okuma ===
file_path = "ORTEST.xlsx"
urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")

# === Veri yapıları ===
urunler = [u for u in urun_kisit_data['Ürün'] if u != "Toplam Maliyet"]
ureticiler = list(set(urun_uretici_data['Üretici']))
satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))

# === Senaryoları oluştur ===
np.random.seed(42)  # Ana seed değeri
sales_scenarios = {}

print("=== Senaryo Oluşturuluyor ===")
for urun in urunler:
    ortalama = urun_param_dict[urun]['ortalama']
    std = urun_param_dict[urun]['std']
    
    # Daha gerçekçi talep dağılımı (simetrik olmayan)
    # Lognormal dağılım kullanımı örneği - negatif değer üretmez
    mu = np.log(ortalama**2 / np.sqrt(ortalama**2 + std**2))
    sigma = np.sqrt(np.log(1 + (std**2 / ortalama**2)))
    
    # Her ürün için SIMULASYON_SAYISI kadar talep senaryosu
    sales_scenarios[urun] = np.random.lognormal(mu, sigma, SIMULASYON_SAYISI)
    
    # Alternatif olarak truncated normal dağılım:
    # sales_scenarios[urun] = np.maximum(0, np.random.normal(ortalama, std, SIMULASYON_SAYISI))

# === Senaryo istatistikleri ===
print("=== Senaryo İstatistikleri ===")
for urun in urunler[:3]:
    print(f"{urun}: Ortalama={np.mean(sales_scenarios[urun]):.2f}, Min={np.min(sales_scenarios[urun]):.2f}, Max={np.max(sales_scenarios[urun]):.2f}")
    print(f"    İlk 5 senaryo: {[round(s, 2) for s in sales_scenarios[urun][:5]]}")

# === Two-stage stochastic model kurulumu ===
solver = pywraplp.Solver.CreateSolver('SCIP')

# --- Birinci Aşama: Üretim karar değişkenleri ---
x = {
    (u, j): solver.IntVar(0, solver.infinity(), f"x_{u}_{j}")
    for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict
}

# --- İkinci Aşama: Satılabilen miktar değişkenleri (senaryo bazlı) ---
# İkinci aşama değişkenlerinin tamsayı olup olmaması probleme bağlıdır
satilan = {
    (u, k): solver.NumVar(0, solver.infinity(), f"satilan_{u}_{k}")
    for u in urunler for k in range(SIMULASYON_SAYISI)
}

# --- Amaç fonksiyonu: Beklenen karı maksimize et ---
objective = solver.Objective()

# Birinci aşama maliyetleri (üretim maliyeti)
birinci_asama_maliyet = sum(x[(u, j)] * urun_uretici_dict[(u, j)] for (u, j) in x)

# İkinci aşama faydası (senaryo bazlı satış geliri)
ikinci_asama_gelir = sum(satilan[(u, k)] * satis_fiyat[u] * SCENARIO_PROB 
                        for u in urunler for k in range(SIMULASYON_SAYISI))

# Doğrudan değişkenlerin katsayılarını ayarla
for (u, j), var in x.items():
    objective.SetCoefficient(var, -urun_uretici_dict[(u, j)])

for u in urunler:
    for k in range(SIMULASYON_SAYISI):
        objective.SetCoefficient(satilan[(u, k)], satis_fiyat[u] * SCENARIO_PROB)

objective.SetMaximization()

# --- Kısıtlar ---
print("\n=== Kısıtlar Oluşturuluyor ===")

# 1. Satış kısıtları: her senaryo için satışlar üretim ve talebi aşamaz
with tqdm(total=len(urunler), desc="Ürün-Senaryo Kısıtları") as progress_bar:
    for u in urunler:
        # Her ürün için toplam üretim miktarı
        toplam_uretim = sum(x[(u, j)] for j in ureticiler if (u, j) in x)
        
        # Her senaryo için satış kısıtları
        for k in range(SIMULASYON_SAYISI):
            # Satış ≤ Üretim
            solver.Add(satilan[(u, k)] <= toplam_uretim)
            
            # Satış ≤ Talep (senaryo k için)
            solver.Add(satilan[(u, k)] <= sales_scenarios[u][k])
        
        progress_bar.update(1)

# 2. Üretici kapasite kısıtları
with tqdm(total=len(ureticiler), desc="Üretici Kapasite Kısıtları") as progress_bar:
    for j in ureticiler:
        toplam_uretim = sum(x[(u, j)] for u in urunler if (u, j) in x)
        
        # Üst kapasite kısıtı
        solver.Add(toplam_uretim <= uretici_kapasite_dict.get(j, float('inf')))
        
        # Alt kapasite kısıtı (minimum üretim miktarı)
        solver.Add(toplam_uretim >= uretici_alt_kapasite_dict.get(j, 0))
        
        progress_bar.update(1)

# === Model bilgileri ===
print("\n=== Model Bilgisi ===")
print(f"Toplam değişken sayısı: {solver.NumVariables():,}")
print(f"- Birinci aşama değişkenleri: {len(x):,}")
print(f"- İkinci aşama değişkenleri: {len(satilan):,}")
print(f"Toplam kısıt sayısı: {solver.NumConstraints():,}")

# === Modeli çöz ===
print("\n=== Model Çözülüyor ===")
start_time = time.time()

with tqdm(total=100, desc="Optimizasyon") as progress_bar:
    progress_bar.update(10)
    status = solver.Solve()
    progress_bar.update(90)

cozum_suresi = time.time() - start_time
print(f"Çözüm süresi: {cozum_suresi:.2f} saniye")

# === Sonuçları raporla ===
if status == pywraplp.Solver.OPTIMAL:
    # Toplam üretim maliyeti (birinci aşama)
    toplam_maliyet = sum(x[(u, j)].solution_value() * urun_uretici_dict[(u, j)] 
                        for (u, j) in x if x[(u, j)].solution_value() > 0)
    
    # Beklenen satış geliri (ikinci aşama)
    beklenen_gelir = sum(satilan[(u, k)].solution_value() * satis_fiyat[u] * SCENARIO_PROB 
                        for u in urunler for k in range(SIMULASYON_SAYISI))
    
    # Beklenen kar
    beklenen_kar = beklenen_gelir - toplam_maliyet
    
    print("\n=== ÇÖZÜM SONUÇLARI ===")
    print(f"Toplam Üretim Maliyeti: {toplam_maliyet:,.2f}")
    print(f"Beklenen Satış Geliri: {beklenen_gelir:,.2f}")
    print(f"Beklenen Kar: {beklenen_kar:,.2f}")
    
    # Üretici bazında sonuçlar
    print("\n=== Üretici Bazında Üretim ===")
    for j in ureticiler:
        toplam_uretici = sum(x[(u, j)].solution_value() for u in urunler if (u, j) in x)
        if toplam_uretici > 0:
            uretici_kullanim = toplam_uretici / uretici_kapasite_dict.get(j, float('inf')) * 100
            print(f"{j} üreticisi:")
            print(f"  - Toplam üretim: {toplam_uretici:.2f}")
            print(f"  - Kapasite kullanımı: {uretici_kullanim:.2f}%")
    
    # Ürün bazında sonuçlar
    print("\n=== Ürün Bazında Üretim ve Satış ===")
    for u in urunler:
        uretilen = sum(x[(u, j)].solution_value() for j in ureticiler if (u, j) in x)
        if uretilen > 0:
            ortalama_satis = sum(satilan[(u, k)].solution_value() for k in range(SIMULASYON_SAYISI)) / SIMULASYON_SAYISI
            ortalama_talep = np.mean(sales_scenarios[u])
            stok_orani = 100 * (1 - ortalama_satis / uretilen) if uretilen > 0 else 0
            
            print(f"{u}:")
            print(f"  - Toplam üretim: {uretilen:.2f}")
            for j in ureticiler:
                if (u, j) in x and x[(u, j)].solution_value() > 0:
                    print(f"    * {j} üreticisi: {x[(u, j)].solution_value():.2f}")
            print(f"  - Ortalama talep: {ortalama_talep:.2f}")
            print(f"  - Ortalama satış: {ortalama_satis:.2f}")
            print(f"  - Ortalama stok oranı: {stok_orani:.2f}%")
    
    # Karşılanamayan talep analizi
    print("\n=== Karşılanamayan Talep Analizi ===")
    for u in urunler:
        uretilen = sum(x[(u, j)].solution_value() for j in ureticiler if (u, j) in x)
        if uretilen > 0:
            karsilanmayan_talepler = [max(0, sales_scenarios[u][k] - satilan[(u, k)].solution_value()) 
                                     for k in range(SIMULASYON_SAYISI)]
            ortalama_eksik = np.mean(karsilanmayan_talepler)
            karsilanmayan_oran = np.mean([1 if k > 0 else 0 for k in karsilanmayan_talepler]) * 100
            
            print(f"{u}:")
            print(f"  - Ortalama karşılanamayan talep: {ortalama_eksik:.2f}")
            print(f"  - Karşılanamayan talep oranı: {karsilanmayan_oran:.2f}%")
            
            # Histogram analizi
            bins = [0, 10, 20, 50, 100, float('inf')]
            hist_values, _ = np.histogram(karsilanmayan_talepler, bins=bins)
            hist_percent = hist_values / SIMULASYON_SAYISI * 100
            
            print(f"  - Karşılanamayan talep dağılımı:")
            for i in range(len(bins)-1):
                print(f"    * {bins[i]}-{bins[i+1] if bins[i+1] != float('inf') else '∞'}: {hist_percent[i]:.2f}%")
else:
    print("Model optimal çözüme ulaşamadı. Durum kodu:", status)
    print("Olası nedenler:")
    if status == pywraplp.Solver.INFEASIBLE:
        print("  - Model kısıtları karşılanamıyor (INFEASIBLE)")
    elif status == pywraplp.Solver.UNBOUNDED:
        print("  - Model sınırsız (UNBOUNDED)")
    else:
        print("  - Diğer bir çözüm hatası")