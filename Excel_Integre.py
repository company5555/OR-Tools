import pandas as pd
from ortools.linear_solver import pywraplp
from statsmodels.tsa.holtwinters import SimpleExpSmoothing

#Ürünlere ait bilgilerin okunması
df = pd.read_excel("ORTEST.xlsx")



#Ürünlerin isimlerinin index olarak belirlenmesi
df.set_index("Ürün", inplace=True)


solver = pywraplp.Solver.CreateSolver("SCIP") #IP çözümler için SCIP, Mixed IP için GLOP
inf = solver.infinity() #Değişkenlerin upper boundunu sonsuz yapmak için kısayol

#Parametreler


TÜretimParametre = df.loc["T-shirt","Üretim Kapasite Sınırı"]
SÜretimParametre = df.loc["Sort","Üretim Kapasite Sınırı"]
KÜretimParametre = df.loc["Kazak","Üretim Kapasite Sınırı"]

TÜrünMaliyetParametre = df.loc["T-shirt","Üretim Maliyeti"]
SÜrünMaliyetParametre = df.loc["Sort","Üretim Maliyeti"]
KÜrünMaliyetParametre = df.loc["Kazak","Üretim Maliyeti"]

ToplamMaliyetParametre = df.loc["Toplam Maliyet","Maliyet Sınırı"]
TToplamMaliyetParametre = df.loc["T-shirt","Maliyet Sınırı"]
SToplamMaliyetParametre = df.loc["Sort","Maliyet Sınırı"]
KToplamMaliyetParametre = df.loc["Kazak","Maliyet Sınırı"]

T_shirtKar = df.loc["T-shirt","Kar"] #Objektif Fonksiyon Katsayıları
SortKar = df.loc["Sort","Kar"]
KazakKar = df.loc["Kazak","Kar"]

T_shirtSatışFiyatı = df.loc["T-shirt","Satış Fiyatı"] #Objektif Fonksiyon Katsayıları
SortSatışFiyatı = df.loc["Sort","Satış Fiyatı"]
KazakSatışFiyatı = df.loc["Kazak","Satış Fiyatı"]


T_shirtSatışOlasılığı = df.loc["T-shirt","Satış Olasılık"] #Objektif Fonksiyon Katsayıları
SortSatışOlasılığı = df.loc["Sort","Satış Olasılık"]
KazakSatışOlasılığı = df.loc["Kazak","Satış Olasılık"]


#Üretim Değişkenleri
T_shirtAdet = solver.IntVar(0, inf,"T_shirtAdet") # TÜretim adında alt sınırı 0, üst sınırı sonsuz olan bir değişken tanımladık.
SortAdet = solver.IntVar(0, inf,"SAdet") # SÜretim adında alt sınırı 0, üst sınırı sonsuz olan bir değişken tanımladık.
KazakAdet = solver.IntVar(0, inf,"KAdet") # KÜretim adında alt sınırı 0, üst sınırı sonsuz olan bir değişken tanımladık.




#Kısıtların belirlenmesi


TÜretimKısıt = solver.Add(T_shirtAdet <= TÜretimParametre)
SÜretimKısıt = solver.Add(SortAdet <= SÜretimParametre)
KÜretimKısıt = solver.Add(KazakAdet <= KÜretimParametre)

MaliyetKısıt = solver.Add(TÜrünMaliyetParametre*T_shirtAdet + SÜrünMaliyetParametre*SortAdet + KÜrünMaliyetParametre*KazakAdet <=ToplamMaliyetParametre)
#Toplam Maliyet
TMaliyetKısıt = solver.Add(TÜrünMaliyetParametre*T_shirtAdet <=TToplamMaliyetParametre) # Ürün Bazında Maliyet
SMaliyetKısıt = solver.Add(SÜrünMaliyetParametre*SortAdet <= SToplamMaliyetParametre)
KMaliyetKısıt = solver.Add(KÜrünMaliyetParametre*KazakAdet <= SToplamMaliyetParametre)


#Objective

solver.Maximize(T_shirtSatışOlasılığı*T_shirtSatışFiyatı*T_shirtAdet
                +SortSatışOlasılığı*SortSatışFiyatı*SortAdet
                +KazakSatışOlasılığı*KazakSatışFiyatı*KazakAdet
                -TÜrünMaliyetParametre*T_shirtAdet
                -SÜrünMaliyetParametre*SortAdet
                -KÜrünMaliyetParametre*KazakAdet)


status = solver.Solve()



if status == pywraplp.Solver.OPTIMAL:
    print(f'Solution:')
    print(f'x = {T_shirtAdet.solution_value():,.2f}'.replace(',', '.'))
    print(f'y = {SortAdet.solution_value():,.2f}'.replace(',', '.'))
    print(f'y = {KazakAdet.solution_value():,.2f}'.replace(',', '.'))
    print(f'Objective value is {solver.Objective().Value():,.2f} dollars.'.replace(',', '.'))
else:
    print('The problem does not have an optimal solution.')









