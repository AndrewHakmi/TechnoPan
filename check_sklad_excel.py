
import pandas as pd
import os

f = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад готовой продукции_спецификация_16.12.2025.xls"

try:
    df = pd.read_excel(f, header=None)
    
    # Search for 598 in any cell
    mask = df.apply(lambda x: x.astype(str).str.contains("598", na=False))
    if mask.any().any():
        print(f"Found '598' in {os.path.basename(f)}:")
        for r_idx in mask[mask.any(axis=1)].index:
            print(f"Row {r_idx}: {df.iloc[r_idx].dropna().tolist()}")
    else:
        print("Value 598 not found in Excel.")

    # Search for 754 (another value from my extraction)
    mask = df.apply(lambda x: x.astype(str).str.contains("754", na=False))
    if mask.any().any():
        print(f"Found '754' in {os.path.basename(f)}:")
        for r_idx in mask[mask.any(axis=1)].index:
            print(f"Row {r_idx}: {df.iloc[r_idx].dropna().tolist()}")
            
except Exception as e:
    print(f"Error: {e}")
