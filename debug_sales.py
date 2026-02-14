
import pandas as pd
import yaml
from utils import load_customer_data
import streamlit as st

# Mock streamlit uploaded file
class MockFile:
    def __init__(self, path):
        self.path = path
        self.name = path.split('\\')[-1]
        
    def read(self):
        with open(self.path, 'rb') as f:
            return f.read()

    def seek(self, pos):
        # Allow seek on the file object logic if needed
        # But load_customer_data calls pd.read_csv(file) which expects a file-like object or path
        # If it's a file-like object, we need to open it.
        pass

# Re-implementing a simple file wrapper because load_customer_data expects something it can seek on if it's not a path
# But pd.read_csv accepts a path string too. However, logic uses file.name.
# Let's pass an open file object.

file_path = '桑原顧客マスタ緯度経度付き_vr1.csv'
with open(file_path, 'rb') as f:
    # Wrap f to have .name
    # But we can't add attribute to built-in file object easily?
    # Actually we can just pass a class that delegates.
    pass

# Easiest: modifying utils to accept path for debugging, or just copy-paste the logic here.
# Let's just run the logic here to see what happens.
    
print("--- Debugging CSV Load ---")

# Try to replicate utils logic
def try_load(path, encodings=['utf-8-sig', 'shift_jis', 'cp932']):
    for enc in encodings:
        try:
            print(f"Trying encoding: {enc}")
            df = pd.read_csv(path, encoding=enc, header=1)
            print("Success!")
            return df, enc
        except Exception as e:
            print(f"Failed with {enc}: {e}")
    return None, None

df, enc = try_load(file_path)

if df is not None:
    print(f"Loaded with {enc}")
    print("Columns:", df.columns.tolist())
    
    # Check config mapping
    with open('config.yaml', 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)
    
    target = config['master_columns']['predicted_sales']
    print(f"Config target for sales: '{target}'")
    
    if target in df.columns:
        print(f"Found '{target}' in columns!")
        print("Sample:", df[target].head())
    else:
        print(f"'{target}' NOT found in columns.")
        # Print hex of columns to see mojibake
        for col in df.columns:
            print(f"'{col}': {col.encode('utf-8')}")

