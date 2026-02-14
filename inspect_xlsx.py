import pandas as pd
import openpyxl

df = pd.read_excel('桑原顧客マスタ緯度経度付き_vr1.xlsx', header=1) # Assuming row 2 is header like CSV
print("Columns:")
for i, col in enumerate(df.columns):
    print(f"{i}: {col}")
