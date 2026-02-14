import pandas as pd

try:
    df = pd.read_csv('桑原顧客マスタ緯度経度付き_vr1.csv', encoding='cp932')
except:
    try:
        df = pd.read_csv('桑原顧客マスタ緯度経度付き_vr1.csv', encoding='utf-8')
    except:
        df = pd.read_csv('桑原顧客マスタ緯度経度付き_vr1.csv', encoding='shift_jis')

print("Columns:")
for i, col in enumerate(df.columns):
    print(f"{i}: {col}")

print("\nSample Data (First 5 rows):")
print(df.iloc[:5, [13, 17, 12, 4]]) # N(13), R(17), M(12), E(4)
