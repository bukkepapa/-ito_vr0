import pandas as pd

try:
    df = pd.read_csv('桑原顧客マスタ緯度経度付き_vr1.csv', encoding='cp932') # shift_jis or cp932 usually for Japanese CSVs
except UnicodeDecodeError:
    try:
        df = pd.read_csv('桑原顧客マスタ緯度経度付き_vr1.csv', encoding='utf-8')
    except:
        df = pd.read_csv('桑原顧客マスタ緯度経度付き_vr1.csv', encoding='shift_jis')

for i, col in enumerate(df.columns):
    print(f"{i}: {col}")
