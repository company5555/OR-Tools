import pandas as pd
import numpy as np
import pulp
from tqdm import tqdm
import time
import threading

# === Parametreler ===
SIMULASYON_SAYISI = 10000  # Belirli sayıda senaryo
np.random.seed(12)

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
urun_alt_kisit_dict = dict(zip(urun_kisit_data['Ürün'], urun_kisit_data['Üretim Alt Sınır']))
urun_ust_kisit_dict = dict(zip(urun_kisit_data['Ürün'], urun_kisit_data['Üretim Üst Sınır']))

# === 1. AŞAMA: Senaryoları oluştur ===
sales_scenarios = {urun: [] for urun in urunler}
for k in range(SIMULASYON_SAYISI):
    np.random.seed(1000 + k)
    for urun in urunler:
        talep = max(0, np.random.normal(urun_param_dict[urun]['ortalama'], urun_param_dict[urun]['std']))
        sales_scenarios[urun].append(talep)

# === 2. AŞAMA: PuLP ile model oluştur ===
print("=== PuLP İLE TÜM INTEGER MODEL ÇÖZÜMÜ ===")

# Model oluştur
model = pulp.LpProblem("Stochastic_Production_Planning", pulp.LpMaximize)

# Karar değişkenleri
# Üretim değişkenleri (integer)
x = pulp.LpVariable.dicts("production", 
                         [(u, j) for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict],
                         lowBound=0, cat='Integer')

# Satış değişkenleri (integer)
y = pulp.LpVariable.dicts("sales",
                         [(u, k) for u in urunler for k in range(SIMULASYON_SAYISI)],
                         lowBound=0, cat='Integer')

# Boolean değişkenler (binary)
b_vars = pulp.LpVariable.dicts("production_decision",
                              urunler,
                              cat='Binary')

# Amaç fonksiyonu: Ortalama karı maksimize et
total_profit = 0
# Üretim maliyetleri (negatif)
for (u, j) in x:
    total_profit -= x[(u, j)] * urun_uretici_dict[(u, j)]

# Satış gelirleri (pozitif, ortalama)
for (u, k) in y:
    total_profit += y[(u, k)] * satis_fiyat[u] / SIMULASYON_SAYISI

model += total_profit, "Total_Profit"

# Kısıtlar
# Üretici kapasite kısıtları
for j in ureticiler:
    production_sum = pulp.lpSum([x[(u, j)] for u in urunler if (u, j) in x])
    model += production_sum <= uretici_kapasite_dict.get(j, float('inf')), f"Max_Capacity_{j}"
    model += production_sum >= uretici_alt_kapasite_dict.get(j, 0), f"Min_Capacity_{j}"

# Ürün alt-üst sınır kısıtları
for u in urunler:
    total_production = pulp.lpSum([x[(u, j)] for j in ureticiler if (u, j) in x])
    alt = urun_alt_kisit_dict[u]
    ust = urun_ust_kisit_dict[u]
    
    # Big-M kısıtları
    M = ust  # Büyük sayı
    model += total_production >= alt * b_vars[u], f"Min_Production_{u}"
    model += total_production <= ust * b_vars[u], f"Max_Production_{u}"

# Satılabilir miktar kısıtları
for u in urunler:
    total_production = pulp.lpSum([x[(u, j)] for j in ureticiler if (u, j) in x])
    for k in range(SIMULASYON_SAYISI):
        model += y[(u, k)] <= total_production, f"Sales_Capacity_{u}_{k}"
        model += y[(u, k)] <= sales_scenarios[u][k], f"Sales_Demand_{u}_{k}"

# Çözüm bilgileri
print("Model çözülüyor... (Tüm değişkenler integer olduğu için uzun sürebilir)")
print(f"Toplam değişken sayısı: {len(model.variables())}")
print(f"Toplam kısıt sayısı: {len(model.constraints)}")
print(f"Integer değişken sayısı: {len(x) + len(y)}")
print(f"Boolean değişken sayısı: {len(b_vars)}")

# Tahmini süre hesaplama
integer_vars = len(x) + len(y)
if integer_vars < 100:
    base_estimate = 0.1 + (integer_vars * 0.01)
elif integer_vars < 1000:
    base_estimate = 1 + (integer_vars * 0.005) + (SIMULASYON_SAYISI * 0.001)
elif integer_vars < 5000:
    base_estimate = 10 + (integer_vars * 0.01) + (SIMULASYON_SAYISI * 0.005)
else:
    base_estimate = 60 + (integer_vars * 0.02) + (SIMULASYON_SAYISI * 0.01)

constraint_factor = len(model.constraints) * 0.0001
complexity_multiplier = 1 + (integer_vars / 10000)
estimated_seconds = (base_estimate + constraint_factor) * complexity_multiplier

min_estimate = max(0.5, estimated_seconds * 0.5)
max_estimate = estimated_seconds * 2.5

print(f"\n📊 TAHMİNİ ÇÖZÜM SÜRESİ:")
print(f"   Problem boyutu        : {'Küçük' if integer_vars < 100 else 'Orta' if integer_vars < 1000 else 'Büyük'}")
print(f"   Integer değişken      : {integer_vars:,}")
print(f"   Minimum beklenen süre : {min_estimate:.1f} saniye")
print(f"   Maksimum beklenen süre : {max_estimate:.1f} saniye")
if max_estimate > 60:
    print(f"   Maksimum beklenen süre : {max_estimate/60:.1f} dakika")

# Progress tracking
solving_finished = False
start_time = time.time()

def show_progress_with_estimate():
    with tqdm(desc="Model çözülüyor", 
              bar_format='{desc}: {elapsed} | {bar} | {postfix}',
              dynamic_ncols=True,
              total=100) as pbar:
        
        while not solving_finished:
            elapsed = time.time() - start_time
            
            if estimated_seconds > 0:
                progress_percent = min(95, (elapsed / estimated_seconds) * 100)
                pbar.n = int(progress_percent)
                pbar.refresh()
            
            if elapsed > 2:
                if estimated_seconds > elapsed:
                    remaining = estimated_seconds - elapsed
                    eta_str = f"ETA: {remaining:.0f}s"
                    if remaining > 60:
                        eta_str = f"ETA: {remaining/60:.1f}dk"
                else:
                    eta_str = "ETA: Hesaplanıyor..."
            else:
                eta_str = f"ETA: ~{estimated_seconds:.0f}s"
            
            pbar.set_postfix_str(f"Geçen: {elapsed:.1f}s | {eta_str}")
            time.sleep(0.5)
        
        final_elapsed = time.time() - start_time
        pbar.n = 100
        pbar.refresh()
        pbar.set_postfix_str(f"Tamamlandı! Süre: {final_elapsed:.1f}s")

print(f"\n🚀 Çözüm başlatılıyor...")

# Progress bar thread'ini başlat
progress_thread = threading.Thread(target=show_progress_with_estimate, daemon=True)
progress_thread.start()

# PuLP ile çöz - CBC solver kullan (ücretsiz ve güçlü)
model.solve(pulp.PULP_CBC_CMD(msg=0))
solving_finished = True
end_time = time.time()

progress_thread.join(timeout=2)

# Çözüm durumu
print(f"\n{'='*70}")
status = pulp.LpStatus[model.status]

if model.status == pulp.LpStatusOptimal:
    print(f"✅ OPTIMAL ÇÖZÜM BULUNDU!")
elif model.status == pulp.LpStatusFeasible:
    print(f"⚠️  UYGUN ÇÖZÜM BULUNDU (optimal olmayabilir)")
elif model.status == pulp.LpStatusInfeasible:
    print(f"❌ UYGUN ÇÖZÜM BULUNAMADI")
elif model.status == pulp.LpStatusUnbounded:
    print(f"❌ PROBLEM SINIRSIZ")
else:
    print(f"❓ BİLİNMEYEN DURUM: {status}")

actual_time = end_time - start_time
print(f"Gerçek çözüm süresi: {actual_time:.2f} saniye")

# Tahmin kalitesi
if estimated_seconds > 0:
    accuracy_diff = abs(actual_time - estimated_seconds)
    accuracy_percent = (accuracy_diff / estimated_seconds) * 100
    
    print(f"Tahmini süre       : {estimated_seconds:.1f} saniye")
    print(f"Tahmin farkı       : {accuracy_diff:.1f} saniye")
    
    if accuracy_percent < 20:
        print(f"🎯 Tahmin kalitesi : Mükemmel (±{accuracy_percent:.0f}%)")
    elif accuracy_percent < 50:
        print(f"📊 Tahmin kalitesi : İyi (±{accuracy_percent:.0f}%)")
    elif accuracy_percent < 100:
        print(f"📈 Tahmin kalitesi : Orta (±{accuracy_percent:.0f}%)")
    else:
        print(f"❓ Tahmin kalitesi : Zayıf (±{accuracy_percent:.0f}%)")

print(f"{'='*70}")

# Sonuçları al
if model.status == pulp.LpStatusOptimal or model.status == pulp.LpStatusFeasible:
    x_values = {k: v.varValue for k, v in x.items()}
    y_values = {k: v.varValue for k, v in y.items()}
    b_values = {k: v.varValue for k, v in b_vars.items()}
    
    # === 3. AŞAMA: Performans metriklerini hesapla ===
    uretim_maliyet = sum(x_values[(u, j)] * urun_uretici_dict[(u, j)] for (u, j) in x_values)
    toplam_gelir = 0
    toplam_stok = 0
    toplam_eksik = 0
    
    for k in range(SIMULASYON_SAYISI):
        for u in urunler:
            satis = y_values[(u, k)]
            talep = sales_scenarios[u][k]
            uretim = sum(x_values[(u, j)] for j in ureticiler if (u, j) in x_values)
            toplam_gelir += satis * satis_fiyat[u]
            toplam_stok += max(0, uretim - talep)
            toplam_eksik += max(0, talep - uretim)

    ortalama_kar = (toplam_gelir - uretim_maliyet) / SIMULASYON_SAYISI
    ortalama_stok = toplam_stok / SIMULASYON_SAYISI
    ortalama_eksik = toplam_eksik / SIMULASYON_SAYISI

    # === SONUÇLAR ===
    print("\n" + "="*50)
    print("SONUÇLAR - PuLP TÜM INTEGER MODEL")
    print("="*50)

    print("\n=== ÜRÜN ÜRETİM KARARLARI ===")
    toplam_uretim_adet = 0
    for (u, j), val in x_values.items():
        if val > 0:
            print(f"{u:15} - {j:10}: {int(val):>6} adet")
            toplam_uretim_adet += int(val)

    print(f"\nToplam üretim adedi: {toplam_uretim_adet:,}")

    print("\n=== ÜRETİLEN ÜRÜNLER ===")
    uretilen_urunler = []
    for u, b_val in b_values.items():
        if b_val > 0.5:
            toplam_uretim = sum(int(x_values[(u, j)]) for j in ureticiler if (u, j) in x_values)
            uretilen_urunler.append(u)
            print(f"✓ {u:15}: {toplam_uretim:>6} adet (Üretiliyor)")
        else:
            print(f"✗ {u:15}: {0:>6} adet (Üretilmiyor)")

    print(f"\nÜretilen ürün çeşidi: {len(uretilen_urunler)}/{len(urunler)}")

    print("\n=== PERFORMANS METRİKLERİ ===")
    print(f"Toplam Üretim Maliyeti     : {uretim_maliyet:>15,.2f} TL")
    print(f"Toplam Gelir               : {toplam_gelir:>15,.2f} TL")
    print(f"Toplam Net Kar             : {toplam_gelir - uretim_maliyet:>15,.2f} TL")
    print(f"Ortalama Senaryo Karı      : {ortalama_kar:>15,.2f} TL")
    print(f"Ortalama Stok Adedi        : {ortalama_stok:>15,.0f} adet")
    print(f"Ortalama Eksik Ürün Adedi  : {ortalama_eksik:>15,.0f} adet")
    
    print(f"Optimal Amaç Fonksiyonu    : {pulp.value(model.objective):>15,.2f} TL")

    print(f"\n{'='*50}")
    print("MODEL ÇÖZÜM BİLGİLERİ")
    print(f"{'='*50}")
    print(f"Çözüm Süresi              : {end_time - start_time:>10.2f} saniye")
    print(f"Çözüm Kalitesi            : {status}")
    print(f"Solver                    : CBC (PuLP)")
    print(f"Toplam Değişken Sayısı     : {len(model.variables()):>10,}")
    print(f"Integer Değişken Sayısı    : {len(x) + len(y):>10,}")
    print(f"Boolean Değişken Sayısı    : {len(b_vars):>10,}")
    print(f"Toplam Kısıt Sayısı        : {len(model.constraints):>10,}")
    print(f"Senaryo Sayısı             : {SIMULASYON_SAYISI:>10,}")

else:
    print("Çözüm bulunamadığı için sonuçlar gösterilemiyor.")

print(f"\n{'='*70}")
print("PuLP KURULUM BİLGİSİ:")
print("pip install pulp")
print(f"{'='*70}")