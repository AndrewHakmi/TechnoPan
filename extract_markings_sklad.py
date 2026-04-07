
import ezdxf
import os
from ezdxf.addons import odafc
import re
import pandas as pd

# Configure ODAFC
odafc.win_exec_path = r"C:\Program Files\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe"

DWG_PATH = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад готовой продукции_раскладка панелей_15.12.2025.dwg"
OUTPUT_FILE = r"c:\Users\freak\TRAE\TechnoPan\sklad_markings_report.txt"

def extract_markings():
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
    # Regex for markings: Starts with п/с/т (case insensitive), optional space/dash, then 2-4 digits
    # Examples: "п 612", "с-598", "Т 100"
    pattern = re.compile(r'\b([пстc])\s*[-]?\s*(\d{2,4})\b', re.IGNORECASE)
    
    counts = {}
    details = [] # Store (layer, text, match)
    
    for e in msp:
        if e.dxftype() in ('TEXT', 'MTEXT'):
            text = e.dxf.text.strip()
            # Clean up formatting codes in MTEXT if any (simple approach)
            text_clean = re.sub(r'\\P', '\n', text) 
            
            matches = pattern.findall(text_clean)
            for prefix, number in matches:
                # Normalize key: "п 612" (lowercase prefix, space, number)
                key = f"{prefix.lower()} {number}"
                counts[key] = counts.get(key, 0) + 1
                details.append({
                    'layer': e.dxf.layer,
                    'original': text,
                    'normalized': key
                })
                
    # 4. Write Report
    print(f"Found {len(counts)} unique marking types.")
    
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(f"Отчет по маркировкам из файла:\n{os.path.basename(DWG_PATH)}\n")
        f.write("="*50 + "\n\n")
        
        f.write("СВОДНАЯ ТАБЛИЦА:\n")
        f.write(f"{'Маркировка':<15} | {'Количество':<10}\n")
        f.write("-" * 28 + "\n")
        
        # Sort by prefix then number (descending length usually implies main panels)
        sorted_keys = sorted(counts.keys(), key=lambda x: (x.split()[0], -int(x.split()[1])))
        
        total_panels = 0
        for k in sorted_keys:
            count = counts[k]
            total_panels += count
            f.write(f"{k:<15} | {count:<10}\n")
            
        f.write("-" * 28 + "\n")
        f.write(f"{'ИТОГО':<15} | {total_panels:<10}\n\n")
        
        f.write("ДЕТАЛИЗАЦИЯ (по слоям):\n")
        df = pd.DataFrame(details)
        if not df.empty:
            summary_layer = df.groupby(['layer', 'normalized']).size().reset_index(name='count')
            f.write(summary_layer.to_string(index=False))
        else:
            f.write("Маркировки не найдены.")

    print(f"Report saved to {OUTPUT_FILE}")
    
    # Print content to console for verification
    with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
        print(f.read())

if __name__ == "__main__":
    extract_markings()
