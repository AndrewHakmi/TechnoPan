
import ezdxf
import os
from ezdxf.addons import odafc
import re
import pandas as pd

# Configure ODAFC
odafc.win_exec_path = r"C:\Program Files\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe"

DWG_PATH = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад готовой продукции_раскладка панелей_15.12.2025.dwg"
OUTPUT_EXCEL = r"c:\Users\freak\TRAE\TechnoPan\sklad_markings_report.xlsx"

def extract_to_excel():
    print(f"Processing {os.path.basename(DWG_PATH)}...")
    
    # 1. Convert to DXF
    temp_dir = os.path.join(os.path.dirname(DWG_PATH), "temp_dxf")
    os.makedirs(temp_dir, exist_ok=True)
    dxf_name = os.path.basename(DWG_PATH).replace(".dwg", ".dxf")
    dxf_path = os.path.join(temp_dir, dxf_name)
    
    if not os.path.exists(dxf_path):
        print("Converting to DXF...")
        odafc.export_dwg(DWG_PATH, temp_dir, fmt="2018", audits=True)
    
    if not os.path.exists(dxf_path):
        print("Error: DXF conversion failed")
        return

    # 2. Read DXF
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    
    # 3. Extract Text
    pattern = re.compile(r'\b([пстc])\s*[-]?\s*(\d{2,4})\b', re.IGNORECASE)
    
    data = []
    
    for e in msp:
        if e.dxftype() in ('TEXT', 'MTEXT'):
            text = e.dxf.text.strip()
            text_clean = re.sub(r'\\P', '\n', text) 
            
            matches = pattern.findall(text_clean)
            for prefix, number in matches:
                # Determine properties
                length_mm = int(number) * 10
                
                # Logic based on layer and prefix
                layer = e.dxf.layer
                width = 0
                
                # Heuristics observed
                norm_prefix = prefix.lower()
                if "1000" in layer or norm_prefix in ['c', 'с', 't', 'т']:
                    width = 1000
                else:
                    width = 1190 # Default for 'п' on standard layers
                
                mark = f"{norm_prefix} {number}"
                
                data.append({
                    'Маркировка': mark,
                    'Префикс': norm_prefix,
                    'Номер': number,
                    'Длина (мм)': length_mm,
                    'Ширина (мм)': width,
                    'Слой': layer,
                    'Исходный текст': text
                })
                
    # 4. Aggregate
    df = pd.DataFrame(data)
    
    if df.empty:
        print("No markings found.")
        return

    # Group by unique characteristics
    summary = df.groupby(['Маркировка', 'Длина (мм)', 'Ширина (мм)', 'Слой']).size().reset_index(name='Количество')
    
    # Sort
    summary = summary.sort_values(by=['Ширина (мм)', 'Длина (мм)'], ascending=[False, False])
    
    # Calculate Area
    summary['Площадь (м2)'] = (summary['Длина (мм)'] / 1000) * (summary['Ширина (мм)'] / 1000) * summary['Количество']
    
    # Calculate Total Area and Total Count for verification
    total_area = summary['Площадь (м2)'].sum()
    total_count = summary['Количество'].sum()
    
    # 5. Write to Excel
    print(f"Saving to {OUTPUT_EXCEL}...")
    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='Сводная', index=False)
        df.to_excel(writer, sheet_name='Детализация', index=False)
        
        # Add summary sheet
        summary_stats = pd.DataFrame({
            'Параметр': ['Всего панелей', 'Общая площадь (м2)'],
            'Значение': [total_count, total_area]
        })
        summary_stats.to_excel(writer, sheet_name='Итоги', index=False)
        
    print("Done!")
    print(summary.to_string())

if __name__ == "__main__":
    extract_to_excel()
