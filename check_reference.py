import pandas as pd
import os

ref_file = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад_спецификация_04.03.2026.xls"

df = pd.read_excel(ref_file, header=None)

# Find panels section
for i in range(10, 50):
    print(f"Row {i}: {df.iloc[i].values}")
