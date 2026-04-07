
import os
import glob
import pandas as pd
import ezdxf
from ezdxf.addons import odafc
import subprocess
import time

# Configure ODAFC
odafc.win_exec_path = r"C:\Program Files\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe"

SEARCH_DIR = r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)"

def analyze_excel(filepath):
    print(f"\n--- Analyzing Excel: {os.path.basename(filepath)} ---")
    try:
        # Try reading with header=None to see raw structure first
        # Use openpyxl for xlsx and xlrd for xls if needed
        if filepath.endswith('.xls'):
             # Install xlrd if missing or handle it. Pandas read_excel handles it usually if xlrd is present.
             pass
             
        df = pd.read_excel(filepath, header=None, nrows=15)
        
        # Look for header row
        header_idx = None
        for idx, row in df.iterrows():
            row_vals = [str(x).lower() for x in row if pd.notna(x)]
            if any("наименование" in x for x in row_vals) or any("марка" in x for x in row_vals):
                header_idx = idx
                break
        
        if header_idx is not None:
            print(f"Header found at row {header_idx}")
            print("Columns:", df.iloc[header_idx].tolist())
            print("First 3 data rows:")
            print(df.iloc[header_idx+1:header_idx+4].to_string())
        else:
            print("Header not found. First 5 rows:")
            print(df.head().to_string())

    except Exception as e:
        print(f"Error reading Excel: {e}")

def convert_dwg_to_dxf(dwg_path, output_dir):
    basename = os.path.basename(dwg_path)
    dxf_name = os.path.splitext(basename)[0] + ".dxf"
    dxf_path = os.path.join(output_dir, dxf_name)
    
    if os.path.exists(dxf_path):
        return dxf_path
        
    print(f"Converting {basename} to DXF...")
    # Use subprocess to call ODAFC directly if ezdxf wrapper fails or is slow
    cmd = [odafc.win_exec_path, os.path.dirname(dwg_path), output_dir, "ACAD2018", "DXF", "0", "1", basename]
    subprocess.run(cmd, capture_output=True)
    
    return dxf_path

def analyze_dwg(filepath):
    print(f"\n--- Analyzing DWG: {os.path.basename(filepath)} ---")
    try:
        # Create temp dir for DXF
        temp_dir = os.path.join(os.path.dirname(filepath), "temp_dxf")
        os.makedirs(temp_dir, exist_ok=True)
        
        dxf_path = convert_dwg_to_dxf(filepath, temp_dir)
        
        if not os.path.exists(dxf_path):
            print("Conversion failed.")
            return

        doc = ezdxf.readfile(dxf_path)
        msp = doc.modelspace()
        
        # Layer Analysis
        layers = set()
        for layer in doc.layers:
            name = layer.dxf.name
            if not name.startswith("0") and "Defpoints" not in name:
                layers.add(name)
        print(f"Unique Layers ({len(layers)}): {sorted(list(layers))[:10]}...")

        # Entity Counts
        counts = {}
        texts = []
        dims = []
        
        for e in msp:
            etype = e.dxftype()
            counts[etype] = counts.get(etype, 0) + 1
            
            if etype in ('TEXT', 'MTEXT'):
                txt = e.dxf.text.strip() if hasattr(e.dxf, 'text') else ""
                if txt and len(texts) < 10:
                    texts.append(f"[{e.dxf.layer}] {txt}")
            
            if etype == 'DIMENSION':
                if len(dims) < 5:
                    dims.append(f"[{e.dxf.layer}] Style: {e.dxf.dimstyle}")

        print("Entity Counts:", counts)
        print("Sample Texts:", texts)
        print("Sample Dimensions:", dims)

    except Exception as e:
        print(f"Error analyzing DWG: {e}")

def main():
    files = glob.glob(os.path.join(SEARCH_DIR, "*.*"))
    dwg_files = [f for f in files if f.lower().endswith('.dwg')]
    
    for dwg in dwg_files:
        analyze_dwg(dwg)
        
        # Find matching excel
        base = os.path.basename(dwg).split('_')[0]
        xls_matches = [f for f in files if f.lower().endswith(('.xls', '.xlsx')) and base in os.path.basename(f)]
        if xls_matches:
            analyze_excel(xls_matches[0])

if __name__ == "__main__":
    main()
