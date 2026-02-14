
import pandas as pd
from utils import load_customer_data, CONFIG
import yaml

# Mock file for Excel
class MockExcelFile:
    def __init__(self, path):
        self.path = path
        self.name = path.split('\\')[-1]
    
    # pd.read_excel accepts path directly too
    # But load_customer_data expects file object behavior?
    # pd.read_excel(file) works with file path string or file-like object.
    # load_customer_data uses file.name check.
    # If we pass a string path to load_customer_data? No, it expects object with .name
    pass

# Just use the path if we can pass a dummy object that has .name but is also a path?
# pd.read_excel calculates engine based on extension if path provided.
# If file-like object provided, it needs engine.
# load_customer_data passes file directly to pd.read_excel.

# Let's bypass load_customer_data for a bit and see what pd.read_excel gives with header=1.
path = '桑原顧客マスタ緯度経度付き_vr1.xlsx'
print(f"Reading {path}...")
df = pd.read_excel(path, header=1)
print("Columns:", df.columns.tolist())

# Check mapping applied in load_customer_data
col_map = CONFIG['master_columns']
print("Expected Sales Col:", col_map['predicted_sales'])

target = col_map['predicted_sales']
if target in df.columns:
    print(f"Found '{target}'!")
    print("Sample:", df[target].head())
else:
    print(f"'{target}' NOT found!")
    # Print columns
    for col in df.columns:
        print(f"'{col}'")

# Simulate rename
rename_map = {
    col_map['customer_code']: 'code',
    col_map['customer_name']: 'name',
    col_map['predicted_sales']: 'sales',
    col_map['latlng']: 'latlng_raw',
    col_map['address1']: 'address',
    col_map.get('work_minutes', '作業時間'): 'WorkMinutes',
    col_map.get('no_entry_time', '入場不可時間帯'): 'NoEntryTime'
}

df_renamed = df.rename(columns=rename_map)
if 'sales' in df_renamed.columns:
    print("Renamed 'sales' column exists.")
    print(df_renamed['sales'].head())
    print("Sum:", df_renamed['sales'].sum())
else:
    print("Renamed 'sales' column DOES NOT exist.")
