
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


@dataclass(frozen=True)
class PanelBlockRule:
    block_name: str
    panel_type: str | None = None
    width_mm_attr: str | None = None
    length_mm_attr: str | None = None
    thickness_mm_attr: str | None = None
    qty_attr: str | None = None
    ral_out_attr: str | None = None
    metal_out_mm_attr: str | None = None
    profile_out_attr: str | None = None
    coating_out: str | None = None
    ral_in_attr: str | None = None
    metal_in_mm_attr: str | None = None
    profile_in_attr: str | None = None
    coating_in: str | None = None


@dataclass(frozen=True)
class Defaults:
    width_mm: float = 1000
    thickness_mm: float = 200
    qty: float = 1
    panel_type: str = ""
    ral_out: str | None = None
    metal_out_mm: float | None = None
    profile_out: str | None = None
    coating_out: str | None = None
    ral_in: str | None = None
    metal_in_mm: float | None = None
    profile_in: str | None = None
    coating_in: str | None = None


@dataclass(frozen=True)
class DimensionPanelType:
    code: str
    length_mm: float


@dataclass(frozen=True)
class DimensionExtraction:
    enabled: bool = False
    dwg_load_version: str | None = "R12"
    marker_text_regex: str = r"^в\s*(\d+)$"
    marker_layers: tuple[str, ...] = ()       # пусто = все слои
    dimension_layers: tuple[str, ...] = ("_РАЗМЕРЫ",)
    panel_width_mm: float = 1000
    measurement_tolerance_mm: float = 2
    distance_y_weight: float = 1.0
    use_height_dimensions: bool = True
    height_assign_radius_mm: float = 50000
    height_perpendicular_tolerance_deg: float = 20
    height_max_perp_mm: float = 12000
    height_max_along_mm: float = 20000
    height_along_weight: float = 0.05
    max_marker_distance_mm: float | None = None
    include_bboxes: tuple[tuple[float, float, float, float], ...] = ()
    panel_types: tuple[DimensionPanelType, ...] = ()


@dataclass(frozen=True)
class TagExtraction:
    enabled: bool = False
    dwg_load_version: str | None = "R12"
    tag_regex: str = r"\b([пстc])\s*[-]?\s*(\d{2,4})\b"
    prefix_width_map: dict[str, float] | None = None
    default_width_mm: float = 1190
    layers: tuple[str, ...] = ()
    layer_prefixes: tuple[str, ...] = ()   # регистронезависимое совпадение по началу имени слоя
    layer_ral_regex: str | None = None     # regex для извлечения RAL из имени слоя, группа 1 = значение
    exclude_layers: tuple[str, ...] = ()


@dataclass(frozen=True)
class Config:
    panel_blocks: tuple[PanelBlockRule, ...]
    defaults: Defaults
    dimension_extraction: DimensionExtraction
    tag_extraction: TagExtraction


def _as_str(v: Any) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def load_config(path: Path) -> Config:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}

    defaults_raw = data.get("defaults") or {}
    defaults = Defaults(
        width_mm=float(defaults_raw.get("width_mm", 1000)),
        thickness_mm=float(defaults_raw.get("thickness_mm", 200)),
        qty=float(defaults_raw.get("qty", 1)),
        panel_type=str(defaults_raw.get("panel_type", "")).strip(),
        ral_out=_as_str(defaults_raw.get("ral_out")),
        metal_out_mm=(float(defaults_raw["metal_out_mm"]) if defaults_raw.get("metal_out_mm") is not None else None),
        profile_out=_as_str(defaults_raw.get("profile_out")),
        coating_out=_as_str(defaults_raw.get("coating_out")),
        ral_in=_as_str(defaults_raw.get("ral_in")),
        metal_in_mm=(float(defaults_raw["metal_in_mm"]) if defaults_raw.get("metal_in_mm") is not None else None),
        profile_in=_as_str(defaults_raw.get("profile_in")),
        coating_in=_as_str(defaults_raw.get("coating_in")),
    )

    rules = []
    for r in (data.get("panel_blocks") or []):
        rules.append(
            PanelBlockRule(
                block_name=str(r.get("block_name", "")).strip(),
                panel_type=_as_str(r.get("panel_type")),
                width_mm_attr=_as_str(r.get("width_mm_attr")),
                length_mm_attr=_as_str(r.get("length_mm_attr")),
                thickness_mm_attr=_as_str(r.get("thickness_mm_attr")),
                qty_attr=_as_str(r.get("qty_attr")),
                ral_out_attr=_as_str(r.get("ral_out_attr")),
                metal_out_mm_attr=_as_str(r.get("metal_out_mm_attr")),
                profile_out_attr=_as_str(r.get("profile_out_attr")),
                coating_out=_as_str(r.get("coating_out")),
                ral_in_attr=_as_str(r.get("ral_in_attr")),
                metal_in_mm_attr=_as_str(r.get("metal_in_mm_attr")),
                profile_in_attr=_as_str(r.get("profile_in_attr")),
                coating_in=_as_str(r.get("coating_in")),
            )
        )

    rules = [r for r in rules if r.block_name]

    dim_raw = data.get("dimension_extraction") or {}
    dim_panel_types = []
    for pt in (dim_raw.get("panel_types") or []):
        code = _as_str(pt.get("code"))
        length = pt.get("length_mm")
        if code and length is not None:
            dim_panel_types.append(DimensionPanelType(code=code, length_mm=float(length)))

    include_bboxes_raw = dim_raw.get("include_bboxes") or []
    include_bboxes = []
    for b in include_bboxes_raw:
        try:
            minx = float(b["minx"])
            miny = float(b["miny"])
            maxx = float(b["maxx"])
            maxy = float(b["maxy"])
        except Exception:
            continue
        include_bboxes.append((minx, miny, maxx, maxy))

    dim = DimensionExtraction(
        enabled=bool(dim_raw.get("enabled", False)),
        dwg_load_version=_as_str(dim_raw.get("dwg_load_version", "R12")),
        marker_text_regex=str(dim_raw.get("marker_text_regex", r"^в\s*(\d+)$")),
        marker_layers=tuple(str(x) for x in (dim_raw.get("marker_layers") or []) if str(x).strip()),
        dimension_layers=tuple(str(x) for x in (dim_raw.get("dimension_layers") or ["_РАЗМЕРЫ"]) if str(x).strip()),
        panel_width_mm=float(dim_raw.get("panel_width_mm", 1000)),
        measurement_tolerance_mm=float(dim_raw.get("measurement_tolerance_mm", 2)),
        distance_y_weight=float(dim_raw.get("distance_y_weight", 1.0)),
        use_height_dimensions=bool(dim_raw.get("use_height_dimensions", True)),
        height_assign_radius_mm=float(dim_raw.get("height_assign_radius_mm", 50000)),
        height_perpendicular_tolerance_deg=float(dim_raw.get("height_perpendicular_tolerance_deg", 20)),
        height_max_perp_mm=float(dim_raw.get("height_max_perp_mm", 12000)),
        height_max_along_mm=float(dim_raw.get("height_max_along_mm", 80000)),
        height_along_weight=float(dim_raw.get("height_along_weight", 0.02)),
        max_marker_distance_mm=(
            float(dim_raw["max_marker_distance_mm"]) if dim_raw.get("max_marker_distance_mm") is not None else None
        ),
        include_bboxes=tuple(include_bboxes),
        panel_types=tuple(dim_panel_types),
    )

    tag_raw = data.get("tag_extraction") or {}
    tag_prefix_map = tag_raw.get("prefix_width_map") or {}
    # Ensure float values
    tag_prefix_map = {k: float(v) for k, v in tag_prefix_map.items()}
    
    tag = TagExtraction(
        enabled=bool(tag_raw.get("enabled", False)),
        dwg_load_version=_as_str(tag_raw.get("dwg_load_version", "R12")),
        tag_regex=str(tag_raw.get("tag_regex", r"\b([пстc])\s*[-]?\s*(\d{2,4})\b")),
        prefix_width_map=tag_prefix_map,
        default_width_mm=float(tag_raw.get("default_width_mm", 1190)),
        layers=tuple(str(x) for x in (tag_raw.get("layers") or []) if str(x).strip()),
        layer_prefixes=tuple(str(x) for x in (tag_raw.get("layer_prefixes") or []) if str(x).strip()),
        layer_ral_regex=_as_str(tag_raw.get("layer_ral_regex")),
        exclude_layers=tuple(str(x) for x in (tag_raw.get("exclude_layers") or []) if str(x).strip()),
    )

    return Config(
        panel_blocks=tuple(rules), 
        defaults=defaults, 
        dimension_extraction=dim,
        tag_extraction=tag
    )
