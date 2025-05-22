import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
from tqdm import tqdm
import time
import threading
import os
from datetime import datetime
import json

class OptimizationTester:
    def __init__(self):
        self.results_summary = []
        self.detailed_results = {}
        self.all_scenarios = {}
        
    def load_data(self, file_path):
        """Excel dosyasından verileri yükle"""
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
        sales_scenarios = {urun: [] for urun in urunler}
        np.random.seed(seed)
        
        for k in range(simulasyon_sayisi):
            for urun in urunler:
                talep = max(0, np.random.normal(urun_param_dict[urun]['ortalama'], urun_param_dict[urun]['std']))
                sales_scenarios[urun].append(talep)
        
        return sales_scenarios
    
    def estimate_solve_time(self, num_variables, integer_vars, constraints, simulasyon_sayisi):
        """Çözüm süresini tahmin et"""
        if integer_vars < 100:
            base_estimate = 0.1 + (integer_vars * 0.01)
        elif integer_vars < 1000:
            base_estimate = 1 + (integer_vars * 0.005) + (simulasyon_sayisi * 0.001)
        elif integer_vars < 5000:
            base_estimate = 10 + (integer_vars * 0.01) + (simulasyon_sayisi * 0.005)
        else:
            base_estimate = 60 + (integer_vars * 0.02) + (simulasyon_sayisi * 0.01)
        
        constraint_factor = constraints * 0.0001
        complexity_multiplier = 1 + (integer_vars / 10000)
        estimated_seconds = (base_estimate + constraint_factor) * complexity_multiplier
        
        return estimated_seconds
    
    def solve_optimization(self, data, sales_scenarios, simulasyon_sayisi, verbose=True):
        """Optimizasyon problemini çöz"""
        solver = pywraplp.Solver.CreateSolver('SCIP')
        
        # Değişkenler
        x = {(u, j): solver.IntVar(0, solver.infinity(), f"x_{u}_{j}") 
             for u in data['urunler'] for j in data['ureticiler'] if (u, j) in data['urun_uretici_dict']}
        
        y = {(u, k): solver.IntVar(0, solver.infinity(), f"y_{u}_{k}") 
             for u in data['urunler'] for k in range(simulasyon_sayisi)}
        
        b_vars = {u: solver.BoolVar(f"b_{u}") for u in data['urunler']}
        
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
        
        # Çözüm süresini tahmin et
        estimated_time = self.estimate_solve_time(
            solver.NumVariables(), 
            len(x) + len(y), 
            solver.NumConstraints(), 
            simulasyon_sayisi
        )
        
        if verbose:
            print(f"   📊 Değişken sayısı: {solver.NumVariables():,}")
            print(f"   📊 Integer değişken: {len(x) + len(y):,}")
            print(f"   📊 Kısıt sayısı: {solver.NumConstraints():,}")
            print(f"   ⏱️  Tahmini süre: {estimated_time:.1f} saniye")
        
        # Çözüm
        start_time = time.time()
        status = solver.Solve()
        end_time = time.time()
        
        solve_time = end_time - start_time
        
        if status in [pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE]:
            x_values = {k: v.solution_value() for k, v in x.items()}
            y_values = {k: v.solution_value() for k, v in y.items()}
            b_values = {k: v.solution_value() for k, v in b_vars.items()}
            
            # Performans metrikleri
            uretim_maliyet = sum(x_values[(u, j)] * data['urun_uretici_dict'][(u, j)] for (u, j) in x_values)
            toplam_gelir = sum(y_values[(u, k)] * data['satis_fiyat'][u] for (u, k) in y_values)
            ortalama_kar = (toplam_gelir - uretim_maliyet) / simulasyon_sayisi
            
            return {
                'status': 'success',
                'solve_time': solve_time,
                'estimated_time': estimated_time,
                'profit': ortalama_kar,
                'total_cost': uretim_maliyet,
                'total_revenue': toplam_gelir,
                'x_values': x_values,
                'y_values': y_values,
                'b_values': b_values,
                'solver_status': status
            }
        else:
            return {
                'status': 'failed',
                'solve_time': solve_time,
                'estimated_time': estimated_time,
                'solver_status': status
            }
    
    def run_single_test(self, file_path, simulasyon_sayisi, seed, verbose=True):
        """Tek bir test çalıştır"""
        if verbose:
            print(f"🚀 Test başlıyor: {os.path.basename(file_path)} | Senaryo: {simulasyon_sayisi} | Seed: {seed}")
        
        # Veri yükleme
        data = self.load_data(file_path)
        if data is None:
            return None
        
        # Senaryoları oluştur
        sales_scenarios = self.generate_scenarios(
            data['urunler'], 
            data['urun_param_dict'], 
            simulasyon_sayisi, 
            seed
        )
        
        # Senaryoları kaydet
        scenario_key = f"{os.path.basename(file_path)}_{simulasyon_sayisi}_{seed}"
        self.all_scenarios[scenario_key] = sales_scenarios
        
        # Optimizasyonu çöz
        result = self.solve_optimization(data, sales_scenarios, simulasyon_sayisi, verbose)
        
        if result:
            result.update({
                'file': os.path.basename(file_path),
                'scenario_count': simulasyon_sayisi,
                'seed': seed,
                'product_count': len(data['urunler']),
                'producer_count': len(data['ureticiler'])
            })
            
            if verbose:
                if result['status'] == 'success':
                    print(f"   ✅ Başarılı - Süre: {result['solve_time']:.2f}s | Kar: {result['profit']:,.2f} TL")
                else:
                    print(f"   ❌ Başarısız - Süre: {result['solve_time']:.2f}s")
        
        return result
    
    def run_comprehensive_test(self):
        """Kapsamlı test çalıştır"""
        files = ["ORTEST50_IP.xlsx", "ORTEST100_IP.xlsx", "ORTEST200_IP.xlsx"]
        scenario_counts = [100, 250, 500, 1000]
        seeds = [1300, 1400, 1500, 1600, 1700]
        
        print("="*80)
        print("🎯 KAPSAMLI OPTİMİZASYON TEST SİSTEMİ")
        print("="*80)
        print(f"📁 Dosya sayısı: {len(files)}")
        print(f"📊 Senaryo sayıları: {scenario_counts}")
        print(f"🎲 Seed değerleri: {seeds}")
        print(f"🔢 Toplam test sayısı: {len(files) * len(scenario_counts) * len(seeds)}")
        print("="*80)
        
        total_tests = len(files) * len(scenario_counts) * len(seeds)
        current_test = 0
        
        for file_path in files:
            if not os.path.exists(file_path):
                print(f"⚠️  Dosya bulunamadı: {file_path}")
                continue
                
            print(f"\n📁 Dosya işleniyor: {file_path}")
            print("-" * 60)
            
            for scenario_count in scenario_counts:
                print(f"\n📊 Senaryo sayısı: {scenario_count}")
                
                scenario_results = []
                for seed in seeds:
                    current_test += 1
                    print(f"\n[{current_test}/{total_tests}] ", end="")
                    
                    result = self.run_single_test(file_path, scenario_count, seed, verbose=True)
                    if result:
                        self.results_summary.append(result)
                        scenario_results.append(result)
                
                # Senaryo grubu özeti
                if scenario_results:
                    successful_results = [r for r in scenario_results if r['status'] == 'success']
                    if successful_results:
                        avg_time = np.mean([r['solve_time'] for r in successful_results])
                        avg_profit = np.mean([r['profit'] for r in successful_results])
                        std_profit = np.std([r['profit'] for r in successful_results])
                        
                        print(f"\n   📈 Özet - Senaryo {scenario_count}:")
                        print(f"   ⏱️  Ortalama süre: {avg_time:.2f}s")
                        print(f"   💰 Ortalama kar: {avg_profit:,.2f} ± {std_profit:,.2f} TL")
                        print(f"   ✅ Başarı oranı: {len(successful_results)}/{len(scenario_results)}")
        
        print(f"\n{'='*80}")
        print("🎉 TÜM TESTLER TAMAMLANDI!")
        print(f"{'='*80}")
        
        # Sonuçları kaydet
        self.save_results()
        self.print_summary()
    
    def save_results(self):
        """Sonuçları dosyalara kaydet"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 1. Özet sonuçlar (CSV)
        if self.results_summary:
            df_results = pd.DataFrame(self.results_summary)
            results_file = f"optimization_results_{timestamp}.csv"
            df_results.to_csv(results_file, index=False)
            print(f"📊 Sonuçlar kaydedildi: {results_file}")
        
        # 2. Detaylı senaryolar (JSON)
        scenarios_file = f"scenarios_{timestamp}.json"
        with open(scenarios_file, 'w', encoding='utf-8') as f:
            # JSON serializable hale getir
            serializable_scenarios = {}
            for key, scenarios in self.all_scenarios.items():
                serializable_scenarios[key] = {
                    product: [float(val) for val in values] 
                    for product, values in scenarios.items()
                }
            json.dump(serializable_scenarios, f, indent=2, ensure_ascii=False)
        print(f"📝 Senaryolar kaydedildi: {scenarios_file}")
        
        # 3. Detaylı metin raporu
        report_file = f"detailed_report_{timestamp}.txt"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("KAPSAMLI OPTİMİZASYON TEST RAPORU\n")
            f.write("="*80 + "\n")
            f.write(f"Rapor tarihi: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Toplam test sayısı: {len(self.results_summary)}\n\n")
            
            # Dosya bazlı analiz
            for file_name in ["ORTEST50_IP.xlsx", "ORTEST100_IP.xlsx", "ORTEST200_IP.xlsx"]:
                file_results = [r for r in self.results_summary if r['file'] == file_name]
                if file_results:
                    f.write(f"\n{'='*60}\n")
                    f.write(f"DOSYA: {file_name}\n")
                    f.write(f"{'='*60}\n")
                    
                    successful_results = [r for r in file_results if r['status'] == 'success']
                    if successful_results:
                        f.write(f"Başarılı test sayısı: {len(successful_results)}/{len(file_results)}\n")
                        f.write(f"Ortalama çözüm süresi: {np.mean([r['solve_time'] for r in successful_results]):.2f} saniye\n")
                        f.write(f"Ortalama kar: {np.mean([r['profit'] for r in successful_results]):,.2f} TL\n")
                        f.write(f"Kar standart sapması: {np.std([r['profit'] for r in successful_results]):,.2f} TL\n\n")
                        
                        # Senaryo sayısı bazlı detay
                        for scenario_count in [100, 250, 500, 1000]:
                            scenario_results = [r for r in successful_results if r['scenario_count'] == scenario_count]
                            if scenario_results:
                                f.write(f"\n--- Senaryo Sayısı: {scenario_count} ---\n")
                                f.write(f"Test sayısı: {len(scenario_results)}\n")
                                f.write(f"Ortalama süre: {np.mean([r['solve_time'] for r in scenario_results]):.2f}s\n")
                                f.write(f"Ortalama kar: {np.mean([r['profit'] for r in scenario_results]):,.2f} TL\n")
                                
                                for seed in [1300, 1400, 1500, 1600, 1700]:
                                    seed_result = next((r for r in scenario_results if r['seed'] == seed), None)
                                    if seed_result:
                                        f.write(f"  Seed {seed}: {seed_result['solve_time']:.2f}s | {seed_result['profit']:,.2f} TL\n")
        
        print(f"📄 Detaylı rapor kaydedildi: {report_file}")
    
    def print_summary(self):
        """Özet istatistikleri yazdır"""
        if not self.results_summary:
            print("❌ Analiz edilecek sonuç bulunamadı.")
            return
        
        successful_results = [r for r in self.results_summary if r['status'] == 'success']
        
        print(f"\n{'='*80}")
        print("📊 GENEL İSTATİSTİKLER")
        print(f"{'='*80}")
        print(f"Toplam test sayısı        : {len(self.results_summary)}")
        print(f"Başarılı test sayısı      : {len(successful_results)}")
        print(f"Başarı oranı              : {len(successful_results)/len(self.results_summary)*100:.1f}%")
        
        if successful_results:
            solve_times = [r['solve_time'] for r in successful_results]
            profits = [r['profit'] for r in successful_results]
            
            print(f"\n⏱️  ÇÖZÜM SÜRESİ ANALİZİ")
            print(f"Ortalama                  : {np.mean(solve_times):>10.2f} saniye")
            print(f"Medyan                    : {np.median(solve_times):>10.2f} saniye")
            print(f"Minimum                   : {np.min(solve_times):>10.2f} saniye")
            print(f"Maksimum                  : {np.max(solve_times):>10.2f} saniye")
            print(f"Standart sapma            : {np.std(solve_times):>10.2f} saniye")
            
            print(f"\n💰 KAR ANALİZİ")
            print(f"Ortalama                  : {np.mean(profits):>15,.2f} TL")
            print(f"Medyan                    : {np.median(profits):>15,.2f} TL")
            print(f"Minimum                   : {np.min(profits):>15,.2f} TL")
            print(f"Maksimum                  : {np.max(profits):>15,.2f} TL")
            print(f"Standart sapma            : {np.std(profits):>15,.2f} TL")
            
            # Dosya bazlı karşılaştırma
            print(f"\n📁 DOSYA BAZLI KARŞILAŞTIRMA")
            for file_name in ["ORTEST50_IP.xlsx", "ORTEST100_IP.xlsx", "ORTEST200_IP.xlsx"]:
                file_results = [r for r in successful_results if r['file'] == file_name]
                if file_results:
                    avg_time = np.mean([r['solve_time'] for r in file_results])
                    avg_profit = np.mean([r['profit'] for r in file_results])
                    print(f"{file_name:15} : {avg_time:>8.2f}s | {avg_profit:>12,.2f} TL | {len(file_results)} test")
            
            # Senaryo sayısı bazlı karşılaştırma
            print(f"\n📊 SENARYO SAYISI BAZLI KARŞILAŞTIRMA")
            for scenario_count in [100, 250, 500, 1000]:
                scenario_results = [r for r in successful_results if r['scenario_count'] == scenario_count]
                if scenario_results:
                    avg_time = np.mean([r['solve_time'] for r in scenario_results])
                    avg_profit = np.mean([r['profit'] for r in scenario_results])
                    print(f"Senaryo {scenario_count:4d}      : {avg_time:>8.2f}s | {avg_profit:>12,.2f} TL | {len(scenario_results)} test")


def main():
    """Ana fonksiyon"""
    print("🚀 Kapsamlı Optimizasyon Test Sistemi Başlatılıyor...")
    
    # Test sistemini başlat
    tester = OptimizationTester()
    tester.run_comprehensive_test()
    
    print(f"\n{'='*80}")
    print("✅ TÜM İŞLEMLER TAMAMLANDI!")
    print("📁 Oluşturulan dosyalar:")
    print("   - optimization_results_[timestamp].csv (Özet sonuçlar)")
    print("   - scenarios_[timestamp].json (Tüm senaryolar)")
    print("   - detailed_report_[timestamp].txt (Detaylı rapor)")
    print(f"{'='*80}")

if __name__ == "__main__":
    main()