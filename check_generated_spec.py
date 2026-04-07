
import pandas as pd
import os

f = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Магазин в Чулыме_раскладка панелей_01.09.2025_spec.xlsx"

if not os.path.exists(f):
    print("File does not exist")
else:
    try:
        df = pd.read_excel(f)
        print(f"Rows: {len(df)}")
        if len(df) > 0:
            print(df.head().to_string())
        else:
            print("File is empty (no rows)")
    except Exception as e:
        print(f"Error reading file: {e}")
