import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
import time
import os
from datetime import datetime
import json

class ORTEST50_Tester:
    def __init__(self):
        self.results = []
        self.all_scenarios = {}
        
    def load_data(self, file_path="ORTEST50_IP.xlsx"):
        """Excel dosyasından verileri yükle"""
        print(f"📁 Veri yükleniyor: {file_path}")
        
        try:
            urun_kisit_data = pd.read_excel(file_path, sheet_name="Ürün - Kısıt")
            urun_satis_data = pd.read_excel(file_path, sheet_name="Ürün - Fiyat")
            urun_uretici_data = pd.read_excel(file_path, sheet_name="Ürün - Üretici")
            uretici_kapasite_data = pd.read_excel(file_path, sheet_name="Üretici - Kapasite")
            urun_param_df = pd.read_excel(file_path, sheet_name="Ürün - Param")
            
            # Veri yapıları
            urunler = [u for u in urun_kisit_data['Ürün'] if u != "Toplam Maliyet"]
            ureticiler = list(set(urun_uretici_data['Üretici']))
            satis_fiyat = dict(zip(urun_satis_data['Ürün'], urun_satis_data['Satış Fiyatı']))
            urun_uretici_dict = {(row['Ürün'], row['Üretici']): row['Birim Maliyet'] for _, row in urun_uretici_data.iterrows()}
            urun_param_dict = {row["Ürün"]: {"ortalama": row["Ortalama"], "std": row["STD"]} for _, row in urun_param_df.iterrows()}
            uretici_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Üst Kapasite']))
            uretici_alt_kapasite_dict = dict(zip(uretici_kapasite_data['Üretici'], uretici_kapasite_data['Alt Kapasite']))
            urun_alt_kisit_dict = dict(zip(urun_kisit_data['Ürün'], urun_kisit_data['Üretim Alt Sınır']))
            urun_ust_kisit_dict = dict(zip(urun_kisit_data['Ürün'], urun_kisit_data['Üretim Üst Sınır']))
            
            print(f"✅ Veri yüklendi - Ürün: {len(urunler)}, Üretici: {len(ureticiler)}")
            
            return {
                'urunler': urunler,
                'ureticiler': ureticiler,
                'satis_fiyat': satis_fiyat,
                'urun_uretici_dict': urun_uretici_dict,
                'urun_param_dict': urun_param_dict,
                'uretici_kapasite_dict': uretici_kapasite_dict,
                'uretici_alt_kapasite_dict': uretici_alt_kapasite_dict,
                'urun_alt_kisit_dict': urun_alt_kisit_dict,
                'urun_ust_kisit_dict': urun_ust_kisit_dict
            }
        except Exception as e:
            print(f"❌ Veri yükleme hatası: {e}")
            return None
    
    def generate_scenarios(self, urunler, urun_param_dict, simulasyon_sayisi, seed):
        """Senaryoları oluştur"""
        print(f"🎲 Senaryolar oluşturuluyor - Sayı: {simulasyon_sayisi}, Seed: {seed}")
        
        sales_scenarios = {urun: [] for urun in urunler}
        np.random.seed(seed)
        
        for k in range(simulasyon_sayisi):
            for urun in urunler:
                talep = max(0, np.random.normal(urun_param_dict[urun]['ortalama'], urun_param_dict[urun]['std']))
                sales_scenarios[urun].append(talep)
        
        print(f"✅ {simulasyon_sayisi} senaryo oluşturuldu")
        return sales_scenarios
    
    def solve_optimization(self, data, sales_scenarios, simulasyon_sayisi):
        """Optimizasyon problemini çöz"""
        print(f"🚀 Optimizasyon modeli kuruluyor...")
        
        solver = pywraplp.Solver.CreateSolver('SCIP')
        
        # Değişkenler
        x = {(u, j): solver.IntVar(0, solver.infinity(), f"x_{u}_{j}") 
             for u in data['urunler'] for j in data['ureticiler'] if (u, j) in data['urun_uretici_dict']}
        
        y = {(u, k): solver.IntVar(0, solver.infinity(), f"y_{u}_{k}") 
             for u in data['urunler'] for k in range(simulasyon_sayisi)}
        
        b_vars = {u: solver.BoolVar(f"b_{u}") for u in data['urunler']}
        
        print(f"📊 Model istatistikleri:")
        print(f"   - Üretim değişkenleri (x): {len(x):,}")
        print(f"   - Satış değişkenleri (y): {len(y):,}")
        print(f"   - Boolean değişkenler (b): {len(b_vars):,}")
        print(f"   - Toplam değişken: {solver.NumVariables():,}")
        
        # Amaç fonksiyonu
        total_profit = solver.Objective()
        for (u, j), var in x.items():
            total_profit.SetCoefficient(var, -data['urun_uretici_dict'][(u, j)])
        for (u, k), var in y.items():
            total_profit.SetCoefficient(var, data['satis_fiyat'][u] / simulasyon_sayisi)
        total_profit.SetMaximization()
        
        # Kısıtlar
        # Üretici kapasite kısıtları
        for j in data['ureticiler']:
            toplam = sum(x[(u, j)] for u in data['urunler'] if (u, j) in x)
            solver.Add(toplam <= data['uretici_kapasite_dict'].get(j, float('inf')))
            solver.Add(toplam >= data['uretici_alt_kapasite_dict'].get(j, 0))
        
        # Ürün alt-üst sınır kısıtları
        for u in data['urunler']:
            toplam_uretim = sum(x[(u, j)] for j in data['ureticiler'] if (u, j) in x)
            alt = data['urun_alt_kisit_dict'][u]
            ust = data['urun_ust_kisit_dict'][u]
            solver.Add(toplam_uretim >= alt * b_vars[u])
            solver.Add(toplam_uretim <= ust * b_vars[u])
        
        # Satılabilir miktar kısıtları
        for u in data['urunler']:
            toplam_uretim = sum(x[(u, j)] for j in data['ureticiler'] if (u, j) in x)
            for k in range(simulasyon_sayisi):
                solver.Add(y[(u, k)] <= toplam_uretim)
                solver.Add(y[(u, k)] <= sales_scenarios[u][k])
        
        print(f"   - Toplam kısıt: {solver.NumConstraints():,}")
        
        # Çözüm
        print(f"⏱️  Model çözülüyor...")
        start_time = time.time()
        status = solver.Solve()
        end_time = time.time()
        
        solve_time = end_time - start_time
        print(f"✅ Çözüm tamamlandı - Süre: {solve_time:.2f} saniye")
        
        if status in [pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE]:
            print(f"🎯 Çözüm durumu: {'Optimal' if status == pywraplp.Solver.OPTIMAL else 'Uygun'}")
            
            x_values = {k: v.solution_value() for k, v in x.items()}
            y_values = {k: v.solution_value() for k, v in y.items()}
            b_values = {k: v.solution_value() for k, v in b_vars.items()}
            
            # Performans metrikleri
            uretim_maliyet = sum(x_values[(u, j)] * data['urun_uretici_dict'][(u, j)] for (u, j) in x_values)
            toplam_gelir = sum(y_values[(u, k)] * data['satis_fiyat'][u] for (u, k) in y_values)
            ortalama_kar = (toplam_gelir - uretim_maliyet) / simulasyon_sayisi
            
            print(f"💰 Performans metrikleri:")
            print(f"   - Toplam maliyet: {uretim_maliyet:,.2f} TL")
            print(f"   - Toplam gelir: {toplam_gelir:,.2f} TL")
            print(f"   - Ortalama kar: {ortalama_kar:,.2f} TL")
            
            return {
                'status': 'success',
                'solve_time': solve_time,
                'profit': ortalama_kar,
                'total_cost': uretim_maliyet,
                'total_revenue': toplam_gelir,
                'solver_status': 'Optimal' if status == pywraplp.Solver.OPTIMAL else 'Feasible'
            }
        else:
            print(f"❌ Çözüm bulunamadı - Durum: {status}")
            return {
                'status': 'failed',
                'solve_time': solve_time,
                'solver_status': status
            }
    
    def run_test(self):
        """ORTEST50 için test çalıştır"""
        print("="*70)
        print("🧪 ORTEST50 TEST SİSTEMİ")
        print("="*70)
        
        # Test parametreleri
        scenario_counts = [100, 250]  # Küçük test için sadece 2 senaryo sayısı
        seeds = [1300, 1400, 1500]    # Sadece 3 seed
        
        print(f"📊 Test parametreleri:")
        print(f"   - Dosya: ORTEST50_IP.xlsx")
        print(f"   - Senaryo sayıları: {scenario_counts}")
        print(f"   - Seed değerleri: {seeds}")
        print(f"   - Toplam test: {len(scenario_counts) * len(seeds)}")
        print("="*70)
        
        # Veriyi yükle
        data = self.load_data("ORTEST50_IP.xlsx")
        if data is None:
            print("❌ Test durduruluyor - veri yüklenemedi")
            return
        
        test_count = 0
        total_tests = len(scenario_counts) * len(seeds)
        
        for scenario_count in scenario_counts:
            print(f"\n{'='*50}")
            print(f"📊 SENARYO SAYISI: {scenario_count}")
            print(f"{'='*50}")
            
            for seed in seeds:
                test_count += 1
                print(f"\n[Test {test_count}/{total_tests}] Seed: {seed}")
                print("-" * 40)
                
                # Senaryoları oluştur
                sales_scenarios = self.generate_scenarios(
                    data['urunler'], 
                    data['urun_param_dict'], 
                    scenario_count, 
                    seed
                )
                
                # Senaryoları kaydet
                scenario_key = f"ORTEST50_{scenario_count}_{seed}"
                self.all_scenarios[scenario_key] = sales_scenarios
                
                # Optimizasyonu çöz
                result = self.solve_optimization(data, sales_scenarios, scenario_count)
                
                if result:
                    result.update({
                        'scenario_count': scenario_count,
                        'seed': seed,
                        'test_number': test_count
                    })
                    self.results.append(result)
                
                print("-" * 40)
        
        # Sonuçları kaydet ve analiz et
        self.save_results()
        self.analyze_results()
    
    def save_results(self):
        """Sonuçları kaydet"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        print(f"\n💾 Sonuçlar kaydediliyor...")
        
        # 1. Test sonuçları (JSON)
        results_file = f"ortest50_results_{timestamp}.json"
        with open(results_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)
        print(f"   ✅ Test sonuçları: {results_file}")
        
        # 2. Senaryolar (JSON)
        scenarios_file = f"ortest50_scenarios_{timestamp}.json"
        with open(scenarios_file, 'w', encoding='utf-8') as f:
            # JSON serializable hale getir
            serializable_scenarios = {}
            for key, scenarios in self.all_scenarios.items():
                serializable_scenarios[key] = {
                    product: [float(val) for val in values] 
                    for product, values in scenarios.items()
                }
            json.dump(serializable_scenarios, f, indent=2, ensure_ascii=False)
        print(f"   ✅ Senaryolar: {scenarios_file}")
        
        # 3. Metin raporu
        report_file = f"ortest50_report_{timestamp}.txt"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("="*70 + "\n")
            f.write("ORTEST50 TEST RAPORU\n")
            f.write("="*70 + "\n")
            f.write(f"Rapor tarihi: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Toplam test sayısı: {len(self.results)}\n\n")
            
            for i, result in enumerate(self.results, 1):
                f.write(f"--- Test {i} ---\n")
                f.write(f"Senaryo sayısı: {result['scenario_count']}\n")
                f.write(f"Seed: {result['seed']}\n")
                f.write(f"Durum: {result['status']}\n")
                f.write(f"Çözüm süresi: {result['solve_time']:.2f} saniye\n")
                if result['status'] == 'success':
                    f.write(f"Kar: {result['profit']:,.2f} TL\n")
                    f.write(f"Çözüm kalitesi: {result['solver_status']}\n")
                f.write("\n")
        
        print(f"   ✅ Detaylı rapor: {report_file}")
    
    def analyze_results(self):
        """Sonuçları analiz et"""
        print(f"\n📊 SONUÇ ANALİZİ")
        print("="*50)
        
        successful_results = [r for r in self.results if r['status'] == 'success']
        
        if not successful_results:
            print("❌ Başarılı test bulunamadı!")
            return
        
        print(f"✅ Başarılı test sayısı: {len(successful_results)}/{len(self.results)}")
        
        # Çözüm süreleri
        solve_times = [r['solve_time'] for r in successful_results]
        print(f"\n⏱️  ÇÖZÜM SÜRESİ ANALİZİ:")
        print(f"   Ortalama: {np.mean(solve_times):.2f} saniye")
        print(f"   Minimum : {np.min(solve_times):.2f} saniye")
        print(f"   Maksimum: {np.max(solve_times):.2f} saniye")
        
        # Kar analizi
        profits = [r['profit'] for r in successful_results]
        print(f"\n💰 KAR ANALİZİ:")
        print(f"   Ortalama: {np.mean(profits):,.2f} TL")
        print(f"   Minimum : {np.min(profits):,.2f} TL")
        print(f"   Maksimum: {np.max(profits):,.2f} TL")
        print(f"   Std. sapma: {np.std(profits):,.2f} TL")
        
        # Senaryo sayısı bazlı analiz
        print(f"\n📊 SENARYO SAYISI BAZLI KARŞILAŞTIRMA:")
        for scenario_count in [100, 250]:
            scenario_results = [r for r in successful_results if r['scenario_count'] == scenario_count]
            if scenario_results:
                avg_time = np.mean([r['solve_time'] for r in scenario_results])
                avg_profit = np.mean([r['profit'] for r in scenario_results])
                print(f"   Senaryo {scenario_count:3d}: {avg_time:6.2f}s | {avg_profit:10,.2f} TL | {len(scenario_results)} test")
        
        # Seed bazlı analiz
        print(f"\n🎲 SEED BAZLI KARŞILAŞTIRMA:")
        for seed in [1300, 1400, 1500]:
            seed_results = [r for r in successful_results if r['seed'] == seed]
            if seed_results:
                avg_time = np.mean([r['solve_time'] for r in seed_results])
                avg_profit = np.mean([r['profit'] for r in seed_results])
                print(f"   Seed {seed}: {avg_time:6.2f}s | {avg_profit:10,.2f} TL | {len(seed_results)} test")
        
        print(f"\n{'='*50}")
        print("🎉 ANALIZ TAMAMLANDI!")
        print(f"{'='*50}")

def main():
    """Ana fonksiyon"""
    # ORTEST50_IP.xlsx dosyasının varlığını kontrol et
    if not os.path.exists("ORTEST50_IP.xlsx"):
        print("❌ ORTEST50_IP.xlsx dosyası bulunamadı!")
        print("   Dosyayı bu scriptin bulunduğu klasöre koyun.")
        return
    
    print("🚀 ORTEST50 Test Sistemi Başlatılıyor...")
    
    # Test sistemini başlat
    tester = ORTEST50_Tester()
    tester.run_test()
    
    print(f"\n{'='*70}")
    print("✅ TEST TAMAMLANDI!")
    print("📁 Oluşturulan dosyalar:")
    print("   - ortest50_results_[timestamp].json (Test sonuçları)")
    print("   - ortest50_scenarios_[timestamp].json (Tüm senaryolar)")
    print("   - ortest50_report_[timestamp].txt (Detaylı rapor)")
    print(f"{'='*70}")

if __name__ == "__main__":
    main()