
import pandas as pd
import os

files = [
    r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Магазин в Чулыме_спецификация_16.09.2025.xls",
    r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад готовой продукции_спецификация_16.12.2025.xls"
]

for f in files:
    print(f"\nScanning {os.path.basename(f)}...")
    try:
        # Read all rows
        df = pd.read_excel(f, header=None)
        
        # Find header
        header_row = -1
        for i, row in df.iterrows():
            s = row.to_string().lower()
            if "наименование" in s or "марка" in s:
                header_row = i
                print(f"Header found at row {i}:")
                print(row.values)
                break
        
        # Search for specific values
        target_vals = ["612", "598"] if "Магазин" in f else ["334"]
        found = False
        for val in target_vals:
            # Search in all cells
            mask = df.apply(lambda x: x.astype(str).str.contains(val, na=False))
            if mask.any().any():
                print(f"Found '{val}' in the file!")
                found = True
                # Print the row(s)
                for r_idx in mask[mask.any(axis=1)].index:
                    print(f"Row {r_idx}: {df.iloc[r_idx].dropna().tolist()}")
        
        if not found:
            print(f"Could not find values {target_vals}")
            
    except Exception as e:
        print(f"Error: {e}")
