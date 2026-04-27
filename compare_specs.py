import pandas as pd
import os

out_file = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад_раслкдка панелей_23.04.2026_spec.xlsx"
ref_file = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад_спецификация_04.03.2026.xls"

print("--- NEW GENERATED SPEC ---")
df_out = pd.read_excel(out_file)
# the rows with area are at the end, 'Площадь, м2 общая' is column 13
# Let's print the last few rows which contain the 'ИТОГО'
print(df_out.tail(5).to_string())

print("\n--- REFERENCE SPEC ---")
df_ref = pd.read_excel(ref_file, header=None)
for idx, row in df_ref.iterrows():
    row_str = [str(x).lower() for x in row.values]
    if any("итого по панелям 1190 мм" in s for s in row_str):
        print(f"Ref total 1: {row.values}")
    if any("итого по панелям" in s and "9006" in s for s in row_str):
        print(f"Ref total 2: {row.values}")
