
from technopan_spec.config import load_config
from technopan_spec.dxf import extract_panels_from_tags, _load_doc
from pathlib import Path
import re

config_path = Path("configs/text_tags.yml")
dwg_path = Path(r"c:\Users\freak\TRAE\TechnoPan\Для разработчиков ИИ (extract.me)\Магазин в Чулыме_раскладка панелей_01.09.2025.dwg")

print(f"Loading config from {config_path}")
cfg = load_config(config_path)

print(f"Regex: {cfg.tag_extraction.tag_regex}")

print("Loading doc...")
# Manually load doc to inspect layers and text
doc = _load_doc(dwg_path, dwg_version="R12")
msp = doc.modelspace()

print("Scanning TEXT entities...")
texts = []
for e in msp:
    if e.dxftype() in ('TEXT', 'MTEXT'):
        texts.append(e)

print(f"Found {len(texts)} text entities.")
regex = re.compile(cfg.tag_extraction.tag_regex, re.IGNORECASE)

for i, e in enumerate(texts[:50]): # Check first 50
    txt = e.dxf.text
    layer = e.dxf.layer
    match = regex.search(txt)
    print(f"[{i}] Layer='{layer}' Text='{txt}' Match={bool(match)}")
    if match:
        print(f"   -> Groups: {match.groups()}")

print("-" * 20)
print("Running full extraction...")
items = extract_panels_from_tags(dwg_path, cfg)
print(f"Extracted {len(items)} items.")
