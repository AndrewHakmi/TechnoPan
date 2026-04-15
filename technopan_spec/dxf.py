
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
import os
import re
import math
import threading

import ezdxf

from .config import Config, PanelBlockRule
from .odafc_utils import resolve_odafc_win_exec_path

# Callable type alias for progress logging
ProgressCb = Callable[[str], None]


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


def _noop(msg: str) -> None:
    pass


def _load_doc(
    path: Path,
    *,
    dwg_version: str | None = None,
    progress_cb: ProgressCb = _noop,
):
    ext = path.suffix.lower()
    if ext == ".dxf":
        progress_cb(f"Чтение DXF: {path.name}")
        return ezdxf.readfile(path)

    if ext in {".dwg", ".dxb"}:
        from ezdxf.addons import odafc

        win_path = resolve_odafc_win_exec_path()
        progress_cb(f"ODA File Converter: {'найден — ' + win_path if win_path else 'НЕ НАЙДЕН'}")
        if win_path:
            ezdxf.options.set("odafc-addon", "win_exec_path", win_path)

        if not odafc.is_installed():
            raise RuntimeError(
                "Для чтения DWG требуется ODA File Converter.\n"
                "Скачайте: https://www.opendesign.com/guestfiles/oda_file_converter\n"
                "Затем задайте переменную окружения ODA_FILE_CONVERTER_EXE (путь к ODAFileConverter.exe)."
            )

        progress_cb(f"Конвертация DWG → DXF (ODA)… это может занять 30–60 секунд")
        if dwg_version:
            doc = odafc.readfile(path, version=dwg_version)
        else:
            doc = odafc.readfile(path)
        progress_cb("Конвертация завершена.")
        return doc

    raise ValueError(f"Неподдерживаемый формат входного файла: {ext}")


# ---------------------------------------------------------------------------
# Auto-detection
# ---------------------------------------------------------------------------

# Regex for tag-style panel codes: "п 612", "с-498", "т 146", "в 346"
_TAG_RE = re.compile(r'(?:^|(?<=\s))([пстcвПСТCВ])\s*[-]?\s*(\d{2,4})(?=\s|$)', re.UNICODE)

# Regex for dimension-mode markers: "в 346", "в346"
_DIM_MARKER_RE = re.compile(r'^в\s*(\d+)$', re.IGNORECASE | re.UNICODE)

# Typical dimension layer names
_DIM_LAYER_HINTS = {"_размеры", "размеры", "сендвичи_размеры", "dimension"}


@dataclass
class AutoDetectResult:
    recommended_config: str          # filename like "text_tags.yml"
    mode: str                        # "tag", "dimension", "attribute", "unknown"
    summary: str                     # human-readable explanation
    details: list[str]               # per-line diagnostic info
    entity_counts: dict[str, int]
    layers: dict[str, int]           # layer → count
    text_samples: list[str]          # first N text strings found
    tag_count: int
    dim_marker_count: int
    insert_count: int


def auto_detect_config(
    path: Path,
    *,
    progress_cb: ProgressCb = _noop,
    stop_event: threading.Event | None = None,
) -> AutoDetectResult:
    """
    Scan the DXF/DWG file and recommend which config/mode to use.
    Does NOT require a Config object — uses heuristics only.
    """
    progress_cb("Автоопределение режима…")

    doc = _load_doc(path, progress_cb=progress_cb)
    msp = doc.modelspace()

    if stop_event and stop_event.is_set():
        raise InterruptedError("Остановлено пользователем")

    progress_cb("Сканирование сущностей…")

    from collections import Counter
    entity_counts: Counter[str] = Counter()
    layer_counts: Counter[str] = Counter()
    text_samples: list[str] = []
    tag_texts: list[tuple[str, str, str]] = []   # (layer, text, prefix+num)
    dim_marker_texts: list[tuple[str, str]] = []
    insert_blocks: Counter[str] = Counter()
    dim_layers_found: set[str] = set()

    for e in msp:
        if stop_event and stop_event.is_set():
            raise InterruptedError("Остановлено пользователем")

        etype = e.dxftype()
        entity_counts[etype] += 1
        try:
            layer = str(e.dxf.layer)
        except Exception:
            layer = ""
        layer_counts[layer] += 1

        if etype in ("TEXT", "MTEXT"):
            if etype == "MTEXT":
                try:
                    text = e.plain_mtext().strip()
                except Exception:
                    text = str(e.dxf.text).strip()
            else:
                text = str(e.dxf.text).strip()

            if text and len(text_samples) < 20:
                text_samples.append(f"[{layer}] {text!r}")

            # Check for panel tags
            matches = _TAG_RE.findall(text)
            for prefix, number in matches:
                tag_texts.append((layer, text, f"{prefix}-{number}"))

            # Check for dimension markers ("в 346")
            if _DIM_MARKER_RE.match(text):
                dim_marker_texts.append((layer, text))

        elif etype == "DIMENSION":
            layer_lo = layer.lower()
            if any(hint in layer_lo for hint in _DIM_LAYER_HINTS):
                dim_layers_found.add(layer)

        elif etype == "INSERT":
            try:
                insert_blocks[str(e.dxf.name)] += 1
            except Exception:
                pass

    # ── Decision logic ──────────────────────────────────────────────────────
    details: list[str] = []
    details.append(f"Файл: {path.name}")
    details.append(f"Типы сущностей: " + ", ".join(
        f"{t}={n}" for t, n in entity_counts.most_common(8)
    ))
    details.append(f"Слоёв: {len(layer_counts)}, топ-8: " + ", ".join(
        f"{l!r}({n})" for l, n in layer_counts.most_common(8)
    ))
    details.append(f"TEXT-тегов найдено: {len(tag_texts)}")
    details.append(f"Маркеров dimension-режима: {len(dim_marker_texts)}")
    details.append(f"INSERT-блоков: {sum(insert_blocks.values())} уник={len(insert_blocks)}")

    # Sample unique tags
    unique_tags = sorted({t[2] for t in tag_texts})
    if unique_tags:
        sample = unique_tags[:12]
        details.append("Примеры тегов: " + ", ".join(sample))

    # Sample unique dim markers
    if dim_marker_texts:
        details.append("Примеры dimension-маркеров: " + ", ".join(
            t[1] for t in dim_marker_texts[:8]
        ))

    if dim_layers_found:
        details.append("DIMENSION слои: " + ", ".join(sorted(dim_layers_found)))

    # Determine mode
    has_tags = len(tag_texts) >= 3
    has_dim = len(dim_marker_texts) >= 1 and len(dim_layers_found) >= 1
    has_insert_attribs = sum(insert_blocks.values()) >= 5

    if has_tags:
        mode = "tag"
        recommended = "text_tags.yml"
        summary = (
            f"Найдено {len(tag_texts)} текстовых тегов панелей "
            f"({len(unique_tags)} уникальных: {', '.join(unique_tags[:6])}{'…' if len(unique_tags)>6 else ''}). "
            f"Рекомендован режим: text_tags."
        )
    elif has_dim:
        mode = "dimension"
        recommended = "abk_dimensions.yml"
        summary = (
            f"Найдено {len(dim_marker_texts)} маркеров dimension-режима "
            f"на слоях: {', '.join(sorted(dim_layers_found))}. "
            f"Рекомендован режим: dimension."
        )
    elif has_insert_attribs:
        mode = "attribute"
        recommended = "default.yml"
        summary = (
            f"Найдено {sum(insert_blocks.values())} INSERT-блоков. "
            f"Рекомендован режим: attribute (default)."
        )
    else:
        mode = "unknown"
        recommended = "default.yml"
        summary = (
            "Не удалось автоматически определить режим. "
            "Проверьте файл и конфиг вручную."
        )

    details.append(f"→ Рекомендован: {recommended} (режим: {mode})")
    progress_cb(f"Автоопределение: {summary}")

    return AutoDetectResult(
        recommended_config=recommended,
        mode=mode,
        summary=summary,
        details=details,
        entity_counts=dict(entity_counts),
        layers=dict(layer_counts),
        text_samples=text_samples,
        tag_count=len(tag_texts),
        dim_marker_count=len(dim_marker_texts),
        insert_count=sum(insert_blocks.values()),
    )


# ---------------------------------------------------------------------------
# Inspection
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Main extraction dispatcher
# ---------------------------------------------------------------------------

def extract_panels_from_dxf(
    path: Path,
    cfg: Config,
    *,
    progress_cb: ProgressCb = _noop,
    stop_event: threading.Event | None = None,
) -> list[PanelItem]:
    if cfg.tag_extraction.enabled:
        return extract_panels_from_tags(path, cfg, progress_cb=progress_cb, stop_event=stop_event)

    if cfg.dimension_extraction.enabled:
        return extract_panels_from_dimensions(path, cfg, progress_cb=progress_cb, stop_event=stop_event)

    return _extract_panels_attribute(path, cfg, progress_cb=progress_cb, stop_event=stop_event)


def _check_stop(stop_event: threading.Event | None) -> None:
    if stop_event and stop_event.is_set():
        raise InterruptedError("Остановлено пользователем")


# ---------------------------------------------------------------------------
# Tag-based extraction
# ---------------------------------------------------------------------------

def extract_panels_from_tags(
    path: Path,
    cfg: Config,
    *,
    progress_cb: ProgressCb = _noop,
    stop_event: threading.Event | None = None,
) -> list[PanelItem]:
    tag_cfg = cfg.tag_extraction
    regex = re.compile(tag_cfg.tag_regex, re.IGNORECASE)

    doc = _load_doc(path, dwg_version=tag_cfg.dwg_load_version, progress_cb=progress_cb)
    _check_stop(stop_event)

    msp = doc.modelspace()
    items: list[PanelItem] = []

    progress_cb("Поиск текстовых тегов (TEXT/MTEXT)…")
    scanned = 0
    skipped_layer = 0
    no_match = 0
    matched_total = 0

    for e in msp:
        _check_stop(stop_event)
        if e.dxftype() not in ("TEXT", "MTEXT"):
            continue

        scanned += 1
        layer = str(e.dxf.layer)

        if tag_cfg.layers and layer not in tag_cfg.layers:
            skipped_layer += 1
            continue
        if tag_cfg.exclude_layers and layer in tag_cfg.exclude_layers:
            skipped_layer += 1
            continue

        if e.dxftype() == "MTEXT":
            try:
                text = e.plain_mtext().strip()
            except Exception:
                text = str(e.dxf.text).strip()
            # Remove remaining MTEXT formatting codes
            text = re.sub(r'\\\w', '', text)
        else:
            text = str(e.dxf.text).strip()

        matches = regex.findall(text)
        if not matches:
            no_match += 1
            continue

        for prefix, number in matches:
            matched_total += 1
            norm_prefix = prefix.lower()
            length_mm = float(number) * 10.0

            width_mm = tag_cfg.default_width_mm
            if tag_cfg.prefix_width_map and norm_prefix in tag_cfg.prefix_width_map:
                width_mm = tag_cfg.prefix_width_map[norm_prefix]

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

    progress_cb(
        f"TEXT/MTEXT просканировано: {scanned} | "
        f"пропущено (слой): {skipped_layer} | "
        f"без совпадений: {no_match} | "
        f"найдено тегов: {matched_total}"
    )

    if not items:
        progress_cb(
            "⚠ Панелей не найдено. Проверьте:\n"
            f"  • tag_regex в конфиге: {tag_cfg.tag_regex!r}\n"
            f"  • layers в конфиге: {list(tag_cfg.layers) or '(все слои)'}\n"
            f"  • Используйте «Автоопределение» чтобы увидеть слои и примеры текстов в файле"
        )

    return items


# ---------------------------------------------------------------------------
# Attribute-based extraction
# ---------------------------------------------------------------------------

def _extract_panels_attribute(
    path: Path,
    cfg: Config,
    *,
    progress_cb: ProgressCb = _noop,
    stop_event: threading.Event | None = None,
) -> list[PanelItem]:
    doc = _load_doc(path, progress_cb=progress_cb)
    _check_stop(stop_event)

    msp = doc.modelspace()
    items: list[PanelItem] = []
    rule_names = {r.block_name for r in cfg.panel_blocks}

    progress_cb(f"Поиск INSERT-блоков: {sorted(rule_names)}")
    scanned = 0
    matched = 0

    for ins in msp.query("INSERT"):
        _check_stop(stop_event)
        scanned += 1
        block_name = str(ins.dxf.name)
        rule = _find_rule(cfg, block_name)
        if not rule:
            continue

        matched += 1
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

    progress_cb(f"INSERT просканировано: {scanned} | совпадений с конфигом: {matched}")
    items = [i for i in items if i.qty and i.length_mm]

    if not items:
        progress_cb(
            "⚠ Панелей не найдено. Проверьте:\n"
            f"  • block_name в конфиге: {sorted(rule_names)}\n"
            f"  • Используйте «Автоопределение» чтобы увидеть блоки в файле\n"
            f"  • Возможно, нужен другой режим (text_tags или dimension)"
        )

    return items


# ---------------------------------------------------------------------------
# Dimension-based extraction
# ---------------------------------------------------------------------------

def extract_panels_from_dimensions(
    path: Path,
    cfg: Config,
    *,
    progress_cb: ProgressCb = _noop,
    stop_event: threading.Event | None = None,
) -> list[PanelItem]:
    dim_cfg = cfg.dimension_extraction
    if not dim_cfg.panel_types:
        raise RuntimeError(
            "dimension_extraction.enabled=true, но panel_types не задан в конфиге."
        )

    code_to_length = {pt.code: pt.length_mm for pt in dim_cfg.panel_types}
    marker_re = re.compile(dim_cfg.marker_text_regex)

    doc = _load_doc(path, dwg_version=dim_cfg.dwg_load_version, progress_cb=progress_cb)
    _check_stop(stop_event)

    msp = doc.modelspace()

    def _in_include_bbox(x: float, y: float) -> bool:
        if not dim_cfg.include_bboxes:
            return True
        for minx, miny, maxx, maxy in dim_cfg.include_bboxes:
            if minx <= x <= maxx and miny <= y <= maxy:
                return True
        return False

    progress_cb(f"Поиск маркеров TEXT (regex: {dim_cfg.marker_text_regex!r})…")
    markers: list[tuple[float, float, float]] = []
    for t in msp.query("TEXT"):
        _check_stop(stop_event)
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

    progress_cb(f"Маркеров найдено: {len(markers)}")

    tol = float(dim_cfg.measurement_tolerance_mm)
    w = float(dim_cfg.panel_width_mm)
    y_weight = float(dim_cfg.distance_y_weight)

    def _seg_from_dim(d) -> tuple[float, float, float, float] | None:
        if not (d.dxf.hasattr("defpoint2") and d.dxf.hasattr("defpoint3")):
            return None
        p2 = d.dxf.defpoint2
        p3 = d.dxf.defpoint3
        return (float(p3.x), float(p3.y), float(p2.x), float(p2.y))

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
                "Нет маркеров TEXT для dimension-режима.\n"
                f"  • marker_text_regex: {dim_cfg.marker_text_regex!r}\n"
                f"  • panel_types: {[pt.code for pt in dim_cfg.panel_types]}\n"
                f"  • include_bboxes: {dim_cfg.include_bboxes or '(весь чертёж)'}"
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

    # Collect height dimensions
    height_dims: list[tuple[float, tuple[float, float], tuple[float, float], tuple[float, float, float, float]]] = []
    if dim_cfg.use_height_dimensions:
        known_lengths = sorted({float(pt.length_mm) for pt in dim_cfg.panel_types})
        progress_cb(f"Сканирование height-DIMENSION (известные длины: {known_lengths})…")
        for d in msp.query("DIMENSION"):
            _check_stop(stop_event)
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
        progress_cb(f"Height-DIMENSION найдено: {len(height_dims)}")

    def _assign_height_dim(
        run_pt: tuple[float, float], run_dir: tuple[float, float] | None
    ) -> float | None:
        if not height_dims or run_dir is None:
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

    progress_cb(f"Сканирование DIMENSION на слоях: {list(dim_cfg.dimension_layers)}…")
    run_segments: dict[tuple[tuple[int, int], tuple[int, int]], tuple[float, float, tuple[float, float] | None]] = {}
    dim_scanned = 0
    dim_matched = 0

    for d in msp.query("DIMENSION"):
        _check_stop(stop_event)
        if str(d.dxf.layer) not in dim_cfg.dimension_layers:
            continue
        dim_scanned += 1
        n = _dim_to_panels(d)
        if n is None or n <= 0:
            continue
        pt = d.dxf.text_midpoint
        if not _in_include_bbox(float(pt.x), float(pt.y)):
            continue
        seg = _seg_from_dim(d)
        if seg is None:
            continue
        dim_matched += 1
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

    progress_cb(
        f"DIMENSION на целевых слоях: {dim_scanned} | "
        f"распознано как панели: {dim_matched} | "
        f"сегментов (дедуп): {len(run_segments)}"
    )

    counts: dict[float, int] = {}
    for mx, my, run_dir in run_segments.values():
        _check_stop(stop_event)
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

    if not items:
        progress_cb(
            "⚠ Панелей не найдено в dimension-режиме. Проверьте:\n"
            f"  • dimension_layers в конфиге: {list(dim_cfg.dimension_layers)}\n"
            f"  • panel_types: {[pt.code for pt in dim_cfg.panel_types]}\n"
            f"  • include_bboxes: {dim_cfg.include_bboxes or '(весь чертёж)'}\n"
            f"  • marker_text_regex: {dim_cfg.marker_text_regex!r}"
        )

    return items
