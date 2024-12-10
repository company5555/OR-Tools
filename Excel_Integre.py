import pandas as pd

from ortools.linear_solver import pywraplp

df = pd.read_csv("ORToolsTest_CSV.csv")


print(df.head())

#print(df.loc["T-shirt"])
print(df.columns)
