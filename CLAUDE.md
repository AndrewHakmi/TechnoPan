# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

TechnoPan automates extraction of panel specifications from CAD layout files (DXF/DWG) into Excel spreadsheets. It supports three extraction modes: attribute-based (from INSERT blocks with DXF attributes), dimension-based (counting panels via DIMENSION entities on specified layers), and tag-based (regex-matching TEXT/MTEXT entities for panel codes).

## Setup

```powershell
python -m pip install -r requirements.txt
```

Dependencies: `ezdxf`, `openpyxl`, `PyYAML`, `xlrd`, `customtkinter`, `packaging`, `pillow`.

Optional DWG support requires [ODA File Converter](https://www.opendesign.com/guestfiles/oda_file_converter):
```powershell
$env:ODA_FILE_CONVERTER_EXE = "C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe"
```

## CLI Commands

```powershell
# Inspect DXF to discover block names and attribute tags
python -m technopan_spec inspect path\to\project.dxf
python -m technopan_spec inspect path\to\project.dxf --json

# Generate specification Excel
python -m technopan_spec generate path\to\project.dxf -c configs\default.yml -o out\spec.xlsx
python -m technopan_spec generate path\to\project.dxf -c configs\default.yml -o out\spec.xlsx --title "My Title"

# Check ODA File Converter availability
python -m technopan_spec odafc-check
python -m technopan_spec odafc-check --print-setx

# Launch GUI
python -m technopan_spec gui
```

## Architecture

The pipeline is: DXF/DWG file → `dxf.py` (extraction) → `spec.py` (grouping + XLSX) → output file.

### Module responsibilities

- **[technopan_spec/config.py](technopan_spec/config.py)** — Loads YAML config into frozen dataclasses: `Config`, `PanelBlockRule`, `Defaults`, `DimensionExtraction`, `DimensionPanelType`, `TagExtraction`.
- **[technopan_spec/dxf.py](technopan_spec/dxf.py)** — Reads DXF/DWG via `ezdxf`. The entry point `extract_panels_from_dxf` dispatches to one of three strategies based on config flags (tag → dimension → attribute, first enabled wins):
  - `extract_panels_from_tags`: matches `TEXT`/`MTEXT` entities against `tag_regex`, converts numeric suffix to `length_mm` (×10), maps prefix letter to width.
  - `extract_panels_from_dimensions`: iterates `DIMENSION` entities on configured layers, splits each into individual panel segments, assigns `length_mm` via perpendicular height dimensions or nearest `TEXT` marker.
  - attribute mode (default): iterates `INSERT` entities, matches block names to `panel_blocks` rules, reads DXF attribute values.
- **[technopan_spec/spec.py](technopan_spec/spec.py)** — Groups `PanelItem` objects by all 12 properties into `PanelRow`, computes totals, writes formatted `.xlsx` with `openpyxl`.
- **[technopan_spec/odafc_utils.py](technopan_spec/odafc_utils.py)** — Resolves ODA File Converter path from `ODA_FILE_CONVERTER_EXE` env var.
- **[technopan_spec/cli.py](technopan_spec/cli.py)** — Argparse CLI wiring four subcommands: `inspect`, `generate`, `odafc-check`, `gui`.
- **[technopan_spec/gui.py](technopan_spec/gui.py)** — `customtkinter`-based GUI frontend for the generate workflow.

### Config structure (YAML)

Three example configs in [configs/](configs/):
- `default.yml` — attribute-based extraction (`panel_blocks` list, no `dimension_extraction`)
- `abk_dimensions.yml` — dimension-based extraction (`dimension_extraction.enabled: true`, `panel_blocks: []`)
- `text_tags.yml` — tag-based extraction (`tag_extraction.enabled: true`)

Exactly one extraction section should have `enabled: true`; the others are ignored. The `defaults` block provides fallback values for all fields not read from the drawing.

### Dimension mode: spatial tuning parameters

`dimension_extraction` has parameters that are non-obvious to tune:
- `include_bboxes` — restricts processing to rectangular regions of modelspace (list of `{minx, miny, maxx, maxy}` in drawing units). Use `inspect --json` to identify coordinate ranges.
- `panel_types` — maps marker codes (e.g. `"346"`) to `length_mm` values; the marker regex captures group 1 as the code.
- `height_assign_radius_mm`, `height_max_perp_mm`, `height_max_along_mm`, `height_along_weight` — control how perpendicular height dimensions are paired with horizontal run dimensions to assign `length_mm`.
- `distance_y_weight` — scales Y distance when finding the nearest marker; increase to prefer markers at the same Y level.

### Data flow in dimension mode

1. Scan `TEXT` entities matching `marker_text_regex` → build `markers` list with `(length_mm, x, y)`
2. Scan `DIMENSION` entities on `dimension_layers` → split each into individual 1000 mm panel segments (deduped by endpoint coordinates)
3. For each segment midpoint, assign `length_mm` by: (a) nearest perpendicular height dimension with a matching known length, or (b) nearest marker by weighted Euclidean distance
4. Aggregate counts per `length_mm` → emit `PanelItem` list using `defaults` for all other fields
