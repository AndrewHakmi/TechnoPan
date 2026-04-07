
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import os
import re
import math

import ezdxf

from .config import Config, PanelBlockRule
from .odafc_utils import resolve_odafc_win_exec_path


@dataclass(frozen=True)
class PanelItem:
    panel_type: str
    ral_out: str | None
    metal_out_mm: float | None
    profile_out: str | None
    coating_out: str | None
    ral_in: str | None
    metal_in_mm: float | None
    profile_in: str | None
    coating_in: str | None
    length_mm: float
    width_mm: float
    thickness_mm: float
    qty: float


def _coerce_float(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _coerce_str(v: Any) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _read_attr(attrs: dict[str, str], tag: str | None) -> str | None:
    if not tag:
        return None
    return attrs.get(tag)


def _find_rule(cfg: Config, block_name: str) -> PanelBlockRule | None:
    for r in cfg.panel_blocks:
        if r.block_name == block_name:
            return r
    return None


def _load_doc(path: Path, *, dwg_version: str | None = None):
    ext = path.suffix.lower()
    if ext == ".dxf":
        return ezdxf.readfile(path)

    if ext in {".dwg", ".dxb"}:
        from ezdxf.addons import odafc

        win_path = resolve_odafc_win_exec_path()
        if win_path:
            ezdxf.options.set("odafc-addon", "win_exec_path", win_path)

        if not odafc.is_installed():
            raise RuntimeError(
                "Для чтения DWG требуется установленный ODA File Converter. "
                "Установите его и задайте переменную окружения ODA_FILE_CONVERTER_EXE "
                "(путь к ODAFileConverter.exe), либо настройте путь в конфиге ezdxf."
            )

        if dwg_version:
            return odafc.readfile(path, version=dwg_version)
        return odafc.readfile(path)

    raise ValueError(f"Неподдерживаемый формат входного файла: {ext}")


def inspect_dxf(path: Path) -> dict[str, dict[str, Any]]:
    doc = _load_doc(path)
    msp = doc.modelspace()

    result: dict[str, dict[str, Any]] = {}
    for ins in msp.query("INSERT"):
        name = str(ins.dxf.name)
        entry = result.setdefault(name, {"count": 0, "attr_tags": set()})
        entry["count"] += 1
        for a in getattr(ins, "attribs", []):
            entry["attr_tags"].add(str(a.dxf.tag))

    for v in result.values():
        v["attr_tags"] = sorted(v["attr_tags"])
    return result


def extract_panels_from_dxf(path: Path, cfg: Config) -> list[PanelItem]:
    if cfg.tag_extraction.enabled:
        return extract_panels_from_tags(path, cfg)

    if cfg.dimension_extraction.enabled:
        return extract_panels_from_dimensions(path, cfg)

    doc = _load_doc(path)
    msp = doc.modelspace()
    items: list[PanelItem] = []

    for ins in msp.query("INSERT"):
        block_name = str(ins.dxf.name)
        rule = _find_rule(cfg, block_name)
        if not rule:
            continue

        attrs = {str(a.dxf.tag): str(a.dxf.text) for a in getattr(ins, "attribs", [])}
        array_mult = 1.0
        for k in ["row_count", "col_count"]:
            if hasattr(ins.dxf, k):
                array_mult *= float(getattr(ins.dxf, k) or 1)

        qty_attr_val = _coerce_float(_read_attr(attrs, rule.qty_attr))
        qty = (qty_attr_val if qty_attr_val is not None else cfg.defaults.qty) * array_mult

        length_mm = _coerce_float(_read_attr(attrs, rule.length_mm_attr)) or 0.0
        width_mm = _coerce_float(_read_attr(attrs, rule.width_mm_attr)) or cfg.defaults.width_mm
        thickness_mm = _coerce_float(_read_attr(attrs, rule.thickness_mm_attr)) or cfg.defaults.thickness_mm

        panel_type = rule.panel_type or cfg.defaults.panel_type or block_name

        ral_out = _coerce_str(_read_attr(attrs, rule.ral_out_attr))
        ral_in = _coerce_str(_read_attr(attrs, rule.ral_in_attr))
        metal_out_mm = _coerce_float(_read_attr(attrs, rule.metal_out_mm_attr))
        metal_in_mm = _coerce_float(_read_attr(attrs, rule.metal_in_mm_attr))
        profile_out = _coerce_str(_read_attr(attrs, rule.profile_out_attr))
        profile_in = _coerce_str(_read_attr(attrs, rule.profile_in_attr))

        items.append(
            PanelItem(
                panel_type=panel_type,
                ral_out=ral_out,
                metal_out_mm=metal_out_mm,
                profile_out=profile_out,
                coating_out=rule.coating_out,
                ral_in=ral_in,
                metal_in_mm=metal_in_mm,
                profile_in=profile_in,
                coating_in=rule.coating_in,
                length_mm=float(length_mm),
                width_mm=float(width_mm),
                thickness_mm=float(thickness_mm),
                qty=float(qty),
            )
        )

    items = [i for i in items if i.qty and i.length_mm]
    return items


def extract_panels_from_tags(path: Path, cfg: Config) -> list[PanelItem]:
    tag_cfg = cfg.tag_extraction
    regex = re.compile(tag_cfg.tag_regex, re.IGNORECASE)
    
    doc = _load_doc(path, dwg_version=tag_cfg.dwg_load_version)
    msp = doc.modelspace()
    
    items: list[PanelItem] = []
    
    for e in msp:
        if e.dxftype() not in ('TEXT', 'MTEXT'):
            continue
            
        layer = str(e.dxf.layer)
        if tag_cfg.layers and layer not in tag_cfg.layers:
            continue
        if tag_cfg.exclude_layers and layer in tag_cfg.exclude_layers:
            continue
            
        text = str(e.dxf.text).strip()
        # Handle MTEXT formatting if necessary (simple removal of \P)
        text = re.sub(r'\\P', '\n', text)
        
        matches = regex.findall(text)
        for prefix, number in matches:
            norm_prefix = prefix.lower()
            length_mm = float(number) * 10.0
            
            # Determine width
            width_mm = tag_cfg.default_width_mm
            if tag_cfg.prefix_width_map and norm_prefix in tag_cfg.prefix_width_map:
                width_mm = tag_cfg.prefix_width_map[norm_prefix]
            
            # Create item
            items.append(
                PanelItem(
                    panel_type=cfg.defaults.panel_type or f"Tag {norm_prefix}",
                    ral_out=cfg.defaults.ral_out,
                    metal_out_mm=cfg.defaults.metal_out_mm,
                    profile_out=cfg.defaults.profile_out,
                    coating_out=cfg.defaults.coating_out,
                    ral_in=cfg.defaults.ral_in,
                    metal_in_mm=cfg.defaults.metal_in_mm,
                    profile_in=cfg.defaults.profile_in,
                    coating_in=cfg.defaults.coating_in,
                    length_mm=length_mm,
                    width_mm=width_mm,
                    thickness_mm=cfg.defaults.thickness_mm,
                    qty=1.0,
                )
            )
            
    return items


def extract_panels_from_dimensions(path: Path, cfg: Config) -> list[PanelItem]:
    dim_cfg = cfg.dimension_extraction
    if not dim_cfg.panel_types:
        raise RuntimeError("dimension_extraction.enabled=true, но panel_types не задан")

    code_to_length = {pt.code: pt.length_mm for pt in dim_cfg.panel_types}
    marker_re = re.compile(dim_cfg.marker_text_regex)

    doc = _load_doc(path, dwg_version=dim_cfg.dwg_load_version)
    msp = doc.modelspace()

    def _in_include_bbox(x: float, y: float) -> bool:
        if not dim_cfg.include_bboxes:
            return True
        for minx, miny, maxx, maxy in dim_cfg.include_bboxes:
            if minx <= x <= maxx and miny <= y <= maxy:
                return True
        return False

    markers: list[tuple[float, float, float]] = []
    for t in msp.query("TEXT"):
        s = str(t.dxf.text).strip()
        m = marker_re.match(s)
        if not m:
            continue
        code = str(m.group(1)).strip()
        if code not in code_to_length:
            continue
        length_mm = float(code_to_length[code])
        tx = float(t.dxf.insert.x)
        ty = float(t.dxf.insert.y)
        if _in_include_bbox(tx, ty):
            markers.append((length_mm, tx, ty))

    marker_lengths = {m[0] for m in markers}

    tol = float(dim_cfg.measurement_tolerance_mm)
    w = float(dim_cfg.panel_width_mm)
    y_weight = float(dim_cfg.distance_y_weight)

    def _seg_from_dim(d) -> tuple[float, float, float, float] | None:
        if not (d.dxf.hasattr("defpoint2") and d.dxf.hasattr("defpoint3")):
            return None
        p2 = d.dxf.defpoint2
        p3 = d.dxf.defpoint3
        return (float(p3.x), float(p3.y), float(p2.x), float(p2.y))

    def _seg_mid(seg: tuple[float, float, float, float]) -> tuple[float, float]:
        x1, y1, x2, y2 = seg
        return ((x1 + x2) / 2.0, (y1 + y2) / 2.0)

    def _seg_dir(seg: tuple[float, float, float, float]) -> tuple[float, float] | None:
        x1, y1, x2, y2 = seg
        dx = x2 - x1
        dy = y2 - y1
        L = math.hypot(dx, dy)
        if L == 0:
            return None
        return (dx / L, dy / L)

    def _dot(a: tuple[float, float], b: tuple[float, float]) -> float:
        return a[0] * b[0] + a[1] * b[1]

    def _perp_dist_point_to_line(
        pt: tuple[float, float], line_pt: tuple[float, float], line_dir: tuple[float, float]
    ) -> float:
        px, py = pt
        x0, y0 = line_pt
        vx, vy = line_dir
        wx, wy = px - x0, py - y0
        return abs(wx * vy - wy * vx)

    def _along_dist(
        pt: tuple[float, float], line_pt: tuple[float, float], line_dir: tuple[float, float]
    ) -> float:
        px, py = pt
        x0, y0 = line_pt
        vx, vy = line_dir
        wx, wy = px - x0, py - y0
        return wx * vx + wy * vy

    def _dim_to_panels(d) -> int | None:
        txt = str(getattr(d.dxf, "text", "") or "").strip()
        m = re.search(r"1000\s*[xх]\s*(\d+)\s*=", txt)
        if m:
            n = int(m.group(1))
            return n if n > 0 else None
        try:
            meas = float(d.get_measurement())
        except Exception:
            return None
        n = int(round(meas / w))
        if n >= 1 and abs(meas - n * w) <= tol:
            return n
        return None

    def _assign_marker(x: float, y: float) -> float:
        if not markers:
            raise RuntimeError(
                "Не найдено маркеров панелей в TEXT для dimension_extraction. "
                "Проверьте marker_text_regex и panel_types в конфиге."
            )
        best_len = markers[0][0]
        best_d = 1e30
        for length_mm, mx, my in markers:
            dx = x - mx
            dy = (y - my) * y_weight
            d2 = dx * dx + dy * dy
            if d2 < best_d:
                best_d = d2
                best_len = length_mm
        return best_len

    height_dims: list[tuple[float, tuple[float, float], tuple[float, float], tuple[float, float, float, float]]] = []
    if dim_cfg.use_height_dimensions:
        known_lengths = sorted({float(pt.length_mm) for pt in dim_cfg.panel_types})
        for d in msp.query("DIMENSION"):
            if str(d.dxf.layer) not in dim_cfg.dimension_layers:
                continue
            try:
                meas = float(d.get_measurement())
            except Exception:
                continue
            seg = _seg_from_dim(d)
            if seg is None:
                continue
            dir_v = _seg_dir(seg)
            if dir_v is None:
                continue
            matched = None
            for L in known_lengths:
                if abs(meas - L) <= tol:
                    matched = L
                    break
            if matched is None:
                continue
            pt = d.dxf.text_midpoint
            hx = float(pt.x)
            hy = float(pt.y)
            if _in_include_bbox(hx, hy):
                height_dims.append((matched, (hx, hy), dir_v, seg))

    def _assign_height_dim(
        run_pt: tuple[float, float], run_dir: tuple[float, float] | None
    ) -> float | None:
        if not height_dims:
            return None
        if run_dir is None:
            return None
        angle_tol = float(dim_cfg.height_perpendicular_tolerance_deg)
        sin_tol = math.sin(math.radians(angle_tol))
        max_perp = float(dim_cfg.height_max_perp_mm)
        max_along = float(dim_cfg.height_max_along_mm)
        along_w = float(dim_cfg.height_along_weight)

        best_len: float | None = None
        best_score = 1e30
        for length_mm, hpt, hdir, _hseg in height_dims:
            if abs(_dot(run_dir, hdir)) > sin_tol:
                continue
            pd = _perp_dist_point_to_line(hpt, run_pt, run_dir)
            ad = abs(_along_dist(hpt, run_pt, run_dir))
            if pd > max_perp or ad > max_along:
                continue
            score = pd + along_w * ad
            if score < best_score:
                best_score = score
                best_len = length_mm

        return best_len

    def _pt_key(x: float, y: float) -> tuple[int, int]:
        return (int(round(x)), int(round(y)))

    run_segments: dict[tuple[tuple[int, int], tuple[int, int]], tuple[float, float, tuple[float, float] | None]] = {}
    for d in msp.query("DIMENSION"):
        if str(d.dxf.layer) not in dim_cfg.dimension_layers:
            continue
        n = _dim_to_panels(d)
        if n is None or n <= 0:
            continue
        pt = d.dxf.text_midpoint
        if not _in_include_bbox(float(pt.x), float(pt.y)):
            continue
        seg = _seg_from_dim(d)
        if seg is None:
            continue
        run_dir = _seg_dir(seg)
        x1, y1, x2, y2 = seg
        dx = (x2 - x1) / float(n)
        dy = (y2 - y1) / float(n)
        for i in range(int(n)):
            ax = x1 + dx * i
            ay = y1 + dy * i
            bx = x1 + dx * (i + 1)
            by = y1 + dy * (i + 1)
            a = _pt_key(ax, ay)
            b = _pt_key(bx, by)
            key = (a, b) if a <= b else (b, a)
            mx = (ax + bx) / 2.0
            my = (ay + by) / 2.0
            run_segments.setdefault(key, (mx, my, run_dir))

    counts: dict[float, int] = {}
    for mx, my, run_dir in run_segments.values():
        length_mm = None
        if dim_cfg.use_height_dimensions:
            length_mm = _assign_height_dim((mx, my), run_dir)
        if length_mm is None:
            length_mm = _assign_marker(mx, my)

        if dim_cfg.max_marker_distance_mm is not None:
            best_d = 1e30
            for _L, mx2, my2 in markers:
                dx2 = mx - mx2
                dy2 = (my - my2) * y_weight
                d2 = dx2 * dx2 + dy2 * dy2
                if d2 < best_d:
                    best_d = d2
            if math.sqrt(best_d) > float(dim_cfg.max_marker_distance_mm):
                continue

        counts[length_mm] = counts.get(length_mm, 0) + 1

    items: list[PanelItem] = []
    for length_mm, qty in sorted(counts.items()):
        if qty <= 0:
            continue
        items.append(
            PanelItem(
                panel_type=cfg.defaults.panel_type or "",
                ral_out=cfg.defaults.ral_out,
                metal_out_mm=cfg.defaults.metal_out_mm,
                profile_out=cfg.defaults.profile_out,
                coating_out=cfg.defaults.coating_out,
                ral_in=cfg.defaults.ral_in,
                metal_in_mm=cfg.defaults.metal_in_mm,
                profile_in=cfg.defaults.profile_in,
                coating_in=cfg.defaults.coating_in,
                length_mm=float(length_mm),
                width_mm=float(cfg.defaults.width_mm),
                thickness_mm=float(cfg.defaults.thickness_mm),
                qty=float(qty),
            )
        )

    return items
