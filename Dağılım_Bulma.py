import pandas as pd
import numpy as np
import scipy.stats as st

def best_fit_distribution(data):
    distributions = [
        st.norm, st.expon, st.gamma, st.lognorm, st.weibull_min, st.weibull_max,
        st.beta, st.uniform, st.triang, st.pearson3, st.genextreme
    ]
    best_distribution = None
    best_p_value = 0
    best_params = None
    
    for distribution in distributions:
        try:
            params = distribution.fit(data)
            _, p = st.kstest(data, distribution.name, args=params)
            
            if p > best_p_value:  # En iyi p-value'ya sahip dağılımı seç
                best_distribution = distribution
                best_p_value = p
                best_params = params
        except Exception:
            continue
    
    return best_distribution, best_params, best_p_value

def normalize_data(data):
    """Veriyi normalize eder (Min-Max veya Log dönüşümü uygular)."""
    data = np.array(data, dtype=np.float64)  # 📌 Tüm verileri float'a çeviriyoruz
    
    if np.any(data <= 0):
        data = (data - data.min()) / (data.max() - data.min() + 1e-9)  # Min-Max Scaling
    else:
        data = np.log1p(data)  # Log dönüşümü (eğer negatif değer yoksa)
    
    return data


def main():
    file_path = 'Ürün - Adet.xlsx'
    df = pd.read_excel(file_path, header=None)  # Başlık olmadığı için header=None
    
    results = []
    
    for index, row in df.iterrows():
        product_name = row[0]  # Ürün ismi
        sales_data = row[1:].dropna().values  # Satış verileri
        
        if len(sales_data) < 2:
            results.append([product_name, "Yetersiz Veri", "-"])
            continue
        
        normalized_data = normalize_data(sales_data)  # 📌 Normalizasyon
        
        best_dist, best_params, p_value = best_fit_distribution(normalized_data)
        
        if best_dist:
            results.append([product_name, best_dist.name, str(best_params), p_value])
        else:
            results.append([product_name, "Belirlenemedi", "-"])
    
    result_df = pd.DataFrame(results, columns=["Ürün", "En İyi Dağılım", "Parametreler", "p-value"])
    result_df.to_excel('Uygun_Dagilimlar.xlsx', index=False)
    print("Dağılımlar belirlendi ve sonuçlar 'Uygun_Dagilimlar.xlsx' dosyasına kaydedildi.")

main()
