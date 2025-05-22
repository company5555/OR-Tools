import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
from tqdm import tqdm

# === Parametreler ===
SIMULASYON_SAYISI = 500  # Belirli sayıda senaryo


# === Excel'den veri okuma ===
file_path = "ORTEST100_IP.xlsx"
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
    np.random.seed(1300 + k)
    for urun in urunler:
        talep = max(0, np.random.normal(urun_param_dict[urun]['ortalama'], urun_param_dict[urun]['std']))
        sales_scenarios[urun].append(talep)

# === 2. AŞAMA: Karar değişkenlerini senaryo bazlı oluştur ve çöz ===
solver = pywraplp.Solver.CreateSolver('SCIP')

# TÜM DEĞİŞKENLER INTEGER YAPILDI
# Üretim değişkenleri (integer)
x = {(u, j): solver.IntVar(0, solver.infinity(), f"x_{u}_{j}") 
     for u in urunler for j in ureticiler if (u, j) in urun_uretici_dict}

# Satış değişkenleri (integer)
y = {(u, k): solver.IntVar(0, solver.infinity(), f"y_{u}_{k}") 
     for u in urunler for k in range(SIMULASYON_SAYISI)}

# Boolean değişkenler (zaten integer - 0 veya 1)
b_vars = {u: solver.BoolVar(f"b_{u}") for u in urunler}

# Amaç fonksiyonu: Ortalama karı maksimize et
total_profit = solver.Objective()
for (u, j), var in x.items():
    total_profit.SetCoefficient(var, -urun_uretici_dict[(u, j)])
for (u, k), var in y.items():
    total_profit.SetCoefficient(var, satis_fiyat[u] / SIMULASYON_SAYISI)
total_profit.SetMaximization()

# Üretici kapasite kısıtları
for j in ureticiler:
    toplam = sum(x[(u, j)] for u in urunler if (u, j) in x)
    solver.Add(toplam <= uretici_kapasite_dict.get(j, float('inf')))
    solver.Add(toplam >= uretici_alt_kapasite_dict.get(j, 0))

# Ürün alt-üst sınır kısıtları (üretiliyorse min-max arasında olmalı)
for u in urunler:
    toplam_uretim = sum(x[(u, j)] for j in ureticiler if (u, j) in x)
    alt = urun_alt_kisit_dict[u]
    ust = urun_ust_kisit_dict[u]
    solver.Add(toplam_uretim >= alt * b_vars[u])
    solver.Add(toplam_uretim <= ust * b_vars[u])

# Satılabilir miktar kısıtları: y <= x ve y <= talep_senaryo
for u in urunler:
    toplam_uretim = sum(x[(u, j)] for j in ureticiler if (u, j) in x)
    for k in range(SIMULASYON_SAYISI):
        solver.Add(y[(u, k)] <= toplam_uretim)
        solver.Add(y[(u, k)] <= sales_scenarios[u][k])

# Modeli çöz
print("=== TÜM INTEGER MODEL ÇÖZÜMÜ ===")
print("Model çözülüyor... (Tüm değişkenler integer olduğu için uzun sürebilir)")
print(f"Toplam değişken sayısı: {solver.NumVariables()}")
print(f"Toplam kısıt sayısı: {solver.NumConstraints()}")
print(f"Integer değişken sayısı: {len(x) + len(y)}")
print(f"Boolean değişken sayısı: {len(b_vars)}")

# Tahmini çözüm süresi hesaplama (Düzeltilmiş formül)
total_vars = solver.NumVariables()
integer_vars = len(x) + len(y)
constraints = solver.NumConstraints()

# Daha gerçekçi tahmini formül
# Küçük problemler için çok düşük, büyük problemler için makul tahminler
if integer_vars < 100:
    # Çok küçük problemler - saniyeler
    base_estimate = 0.1 + (integer_vars * 0.01)
elif integer_vars < 1000:
    # Küçük-orta problemler - dakikalar
    base_estimate = 1 + (integer_vars * 0.005) + (SIMULASYON_SAYISI * 0.001)
elif integer_vars < 5000:
    # Büyük problemler - dakikalar/saatler
    base_estimate = 10 + (integer_vars * 0.01) + (SIMULASYON_SAYISI * 0.005)
else:
    # Çok büyük problemler - saatler
    base_estimate = 60 + (integer_vars * 0.02) + (SIMULASYON_SAYISI * 0.01)

# Kısıt faktörü (daha düşük)
constraint_factor = constraints * 0.0001

# Karmaşıklık çarpanı (daha muhafazakar)
complexity_multiplier = 1 + (integer_vars / 10000)

estimated_seconds = (base_estimate + constraint_factor) * complexity_multiplier

# Tahmini süre aralığı (daha dar aralık)
min_estimate = max(0.5, estimated_seconds * 0.5)
max_estimate = estimated_seconds * 2.5

print(f"\n📊 TAHMİNİ ÇÖZÜM SÜRESİ (Düzeltilmiş):")
print(f"   Problem boyutu        : {'Küçük' if integer_vars < 100 else 'Orta' if integer_vars < 1000 else 'Büyük'}")
print(f"   Integer değişken      : {integer_vars:,}")
print(f"   Minimum beklenen süre : {min_estimate:.1f} saniye")
print(f"   Maksimum beklenen süre : {max_estimate:.1f} saniye")
if max_estimate > 60:
    print(f"   Maksimum beklenen süre : {max_estimate/60:.1f} dakika")
print(f"   Ortalama beklenti     : {estimated_seconds:.1f} saniye")

# Uyarı mesajları
if integer_vars > 1000:
    print(f"   ⚠️  Büyük problem - uzun sürebilir")
elif integer_vars < 100:
    print(f"   ✅ Küçük problem - hızlı çözülmeli")
else:
    print(f"   📊 Orta boyut problem - makul süre")
    
print(f"   💡 Bu tahminler deneyimsel, gerçek süre farklı olabilir")

import time
import threading

# Progress bar için global değişkenler
solving_finished = False
start_time = time.time()

# Gelişmiş progress bar fonksiyonu (tahmini süre ile)
def show_progress_with_estimate():
    with tqdm(desc="Model çözülüyor", 
              bar_format='{desc}: {elapsed} | {bar} | {postfix}',
              dynamic_ncols=True,
              total=100) as pbar:
        
        while not solving_finished:
            elapsed = time.time() - start_time
            
            # Tahmini tamamlanma yüzdesi (estimated_seconds'a göre)
            if estimated_seconds > 0:
                progress_percent = min(95, (elapsed / estimated_seconds) * 100)
                pbar.n = int(progress_percent)
                pbar.refresh()
            
            # Kalan süre tahmini
            if elapsed > 2:  # İlk 2 saniye sonra tahmin başlat
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
        
        # Çözüm tamamlandığında son güncelleme
        final_elapsed = time.time() - start_time
        pbar.n = 100
        pbar.refresh()
        
        # Tahmin kalitesi değerlendirmesi
        if estimated_seconds > 0:
            accuracy = abs(final_elapsed - estimated_seconds) / estimated_seconds * 100
            if accuracy < 20:
                accuracy_emoji = "🎯"
            elif accuracy < 50:
                accuracy_emoji = "📊"
            else:
                accuracy_emoji = "❓"
        else:
            accuracy_emoji = "📊"
            
        pbar.set_postfix_str(f"Tamamlandı! Süre: {final_elapsed:.1f}s {accuracy_emoji}")

print(f"\n🚀 Çözüm başlatılıyor...")

# Progress bar thread'ini başlat
progress_thread = threading.Thread(target=show_progress_with_estimate, daemon=True)
progress_thread.start()

# Modeli çöz
status = solver.Solve()
solving_finished = True
end_time = time.time()

# Thread'in bitmesini bekle
progress_thread.join(timeout=2)

# Çözüm durumu kontrolü
print(f"\n{'='*70}")
if status == pywraplp.Solver.OPTIMAL:
    print(f"✅ OPTIMAL ÇÖZÜM BULUNDU!")
elif status == pywraplp.Solver.FEASIBLE:
    print(f"⚠️  UYGUN ÇÖZÜM BULUNDU (optimal olmayabilir)")
elif status == pywraplp.Solver.INFEASIBLE:
    print(f"❌ UYGUN ÇÖZÜM BULUNAMADI")
elif status == pywraplp.Solver.UNBOUNDED:
    print(f"❌ PROBLEM SINIRSIZ")
else:
    print(f"❓ BİLİNMEYEN DURUM: {status}")

actual_time = end_time - start_time
print(f"Gerçek çözüm süresi: {actual_time:.2f} saniye")

# Tahmin kalitesi analizi
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

# Sonuçları al (sadece çözüm bulunduysa)
if status in [pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE]:
    x_values = {k: v.solution_value() for k, v in x.items()}
    y_values = {k: v.solution_value() for k, v in y.items()}
    b_values = {k: v.solution_value() for k, v in b_vars.items()}
else:
    print("Çözüm bulunamadığı için sonuçlar gösterilemiyor.")
    exit()

# === 3. AŞAMA: Performans metriklerini hesapla ===
uretim_maliyet = sum(x_values[(u, j)] * urun_uretici_dict[(u, j)] for (u, j) in x_values)
toplam_gelir = 0
toplam_stok = 0
toplam_eksik = 0
for k in range(SIMULASYON_SAYISI):
    for u in urunler:
        satis = y_values[(u, k)]
        talep = sales_scenarios[u][k]
        uretim = sum(x_values[(u, j)] for j in ureticiler if (u, j) in x)
        toplam_gelir += satis * satis_fiyat[u]
        toplam_stok += max(0, uretim - talep)
        toplam_eksik += max(0, talep - uretim)

ortalama_kar = (toplam_gelir - uretim_maliyet) / SIMULASYON_SAYISI
ortalama_stok = toplam_stok / SIMULASYON_SAYISI
ortalama_eksik = toplam_eksik / SIMULASYON_SAYISI

# === SONUÇLAR ===
print("\n" + "="*50)
print("SONUÇLAR - TÜM INTEGER MODEL")
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
    if b_val > 0.5:  # Boolean değişken 1 ise
        toplam_uretim = sum(int(x_values[(u, j)]) for j in ureticiler if (u, j) in x_values)
        uretilen_urunler.append(u)
        print(f"✓ {u:15}: {toplam_uretim:>6} adet (Üretiliyor)")
    else:
        print(f"✗ {u:15}: {0:>6} adet (Üretilmiyor)")

print(f"\nÜretilen ürün çeşidi: {len(uretilen_urunler)}/{len(urunler)}")

print("\n=== İLK 5 SENARYO SATIŞ DETAYLARI ===")
for k in range(min(5, SIMULASYON_SAYISI)):
    print(f"\n--- Senaryo {k + 1} ---")
    senaryo_satis_toplam = 0
    for u in urunler:
        talep = int(sales_scenarios[u][k])
        satis = int(y_values[(u, k)])
        uretim = sum(int(x_values[(u, j)]) for j in ureticiler if (u, j) in x_values)
        
        if satis > 0 or talep > 0:
            karsilama_orani = (satis / talep * 100) if talep > 0 else 0
            print(f"{u:12} | Talep:{talep:>4} | Üretim:{uretim:>4} | Satış:{satis:>4} | Karşılama:%{karsilama_orani:>3.0f}")
            senaryo_satis_toplam += satis
    print(f"Senaryo toplam satış: {senaryo_satis_toplam} adet")

print("\n=== PERFORMANS METRİKLERİ ===")
print(f"Toplam Üretim Maliyeti     : {uretim_maliyet:>15,.2f} TL")
print(f"Toplam Gelir               : {toplam_gelir:>15,.2f} TL")
print(f"Toplam Net Kar             : {toplam_gelir - uretim_maliyet:>15,.2f} TL")
print(f"Ortalama Senaryo Karı      : {ortalama_kar:>15,.2f} TL")
print(f"Ortalama Stok Adedi        : {ortalama_stok:>15,.0f} adet")
print(f"Ortalama Eksik Ürün Adedi  : {ortalama_eksik:>15,.0f} adet")

# Ek integer-specific metrikler
stok_orani = (ortalama_stok / toplam_uretim_adet * 100) if toplam_uretim_adet > 0 else 0
print(f"Ortalama Stok Oranı        : {stok_orani:>15.1f}%")

print(f"\n{'='*50}")
print("MODEL ÇÖZÜM BİLGİLERİ")
print(f"{'='*50}")
print(f"Çözüm Süresi              : {end_time - start_time:>10.2f} saniye")
print(f"Çözüm Kalitesi            : {'Optimal' if status == pywraplp.Solver.OPTIMAL else 'Uygun'}")
print(f"Toplam Değişken Sayısı     : {solver.NumVariables():>10,}")
print(f"Integer Değişken Sayısı    : {len(x) + len(y):>10,}")
print(f"Boolean Değişken Sayısı    : {len(b_vars):>10,}")
print(f"Toplam Kısıt Sayısı        : {solver.NumConstraints():>10,}")
print(f"Senaryo Sayısı             : {SIMULASYON_SAYISI:>10,}") 