
import ezdxf
import os
from ezdxf.addons import odafc
import re

# Configure ODAFC
odafc.win_exec_path = r"C:\Program Files\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe"

files = [
    r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Магазин в Чулыме_раскладка панелей_01.09.2025.dwg",
    r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Склад готовой продукции_раскладка панелей_15.12.2025.dwg"
]

def verify_counts(dwg_path, target_texts):
    print(f"\nVerifying {os.path.basename(dwg_path)}...")
    temp_dir = os.path.join(os.path.dirname(dwg_path), "temp_dxf")
    dxf_name = os.path.basename(dwg_path).replace(".dwg", ".dxf")
    dxf_path = os.path.join(temp_dir, dxf_name)
    
    if not os.path.exists(dxf_path):
        print("DXF not found, skipping (run analyze first)")
        return

    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    
    counts = {t: 0 for t in target_texts}
    
    # Regex to match "п 612" or "п-612" allowing flexible spacing
    patterns = {t: re.compile(re.escape(t).replace(r'\ ', r'\s*[-]?\s*'), re.IGNORECASE) for t in target_texts}
    
    for e in msp:
        if e.dxftype() in ('TEXT', 'MTEXT'):
            txt = e.dxf.text.strip()
            # Check against targets
            for target, pattern in patterns.items():
                # We search for the number part specifically if needed, but here we look for exact text match
                # The text in DWG was "п 612"
                if pattern.search(txt):
                    counts[target] += 1
                    
    print("Counts found in DWG:", counts)

if __name__ == "__main__":
    # Targets derived from previous Excel analysis
    # Магазин: п 612 (expect 9), п 598 (expect 14), п 334 (none expected here)
    verify_counts(files[0], ["п 612", "п 598"])
    
    # Склад: п 334 (expect 5)
    verify_counts(files[1], ["п 334"])
