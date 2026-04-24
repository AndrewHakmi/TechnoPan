from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

from .dxf import PanelItem


@dataclass(frozen=True)
class PanelRow:
    idx: int
    supply_no: str
    panel_type: str
    tag_prefix: str | None          # буква маркировки, напр. "п"
    tag_number: int | None          # цифра маркировки, напр. 612
    ral_out: str | None
    metal_out_mm: float | None
    profile_out: str | None
    ral_in: str | None
    metal_in_mm: float | None
    profile_in: str | None
    length_mm: float
    width_mm: float
    thickness_mm: float
    qty: float
    area_m2_total: float
    coating_out: str | None
    coating_in: str | None


def _group_key(i: PanelItem) -> tuple:
    return (
        i.panel_type,
        i.tag_prefix,
        i.tag_number,
        i.ral_out,
        i.metal_out_mm,
        i.profile_out,
        i.coating_out,
        i.ral_in,
        i.metal_in_mm,
        i.profile_in,
        i.coating_in,
        i.length_mm,
        i.width_mm,
        i.thickness_mm,
    )


def build_panel_rows(items: list[PanelItem]) -> list[PanelRow]:
    grouped: dict[tuple, dict[str, float]] = {}
    for i in items:
        key = _group_key(i)
        g = grouped.setdefault(key, {"qty": 0.0, "area": 0.0})
        g["qty"] += float(i.qty)
        g["area"] += (float(i.length_mm) * float(i.width_mm) / 1_000_000.0) * float(i.qty)

    rows: list[PanelRow] = []
    for n, (key, agg) in enumerate(sorted(grouped.items(), key=lambda kv: kv[0]), start=1):
        (
            panel_type,
            tag_prefix,
            tag_number,
            ral_out,
            metal_out_mm,
            profile_out,
            coating_out,
            ral_in,
            metal_in_mm,
            profile_in,
            coating_in,
            length_mm,
            width_mm,
            thickness_mm,
        ) = key

        rows.append(
            PanelRow(
                idx=n,
                supply_no="ПК-",
                panel_type=str(panel_type),
                tag_prefix=tag_prefix,
                tag_number=tag_number,
                ral_out=ral_out,
                metal_out_mm=metal_out_mm,
                profile_out=profile_out,
                ral_in=ral_in,
                metal_in_mm=metal_in_mm,
                profile_in=profile_in,
                length_mm=float(length_mm),
                width_mm=float(width_mm),
                thickness_mm=float(thickness_mm),
                qty=round(float(agg["qty"]), 3),
                area_m2_total=round(float(agg["area"]), 3),
                coating_out=coating_out,
                coating_in=coating_in,
            )
        )
    return rows


def write_spec_xlsx(path: Path, rows: list[PanelRow], title: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Панели"

    headers = [
        "№ п.п.",
        "№ поставки",
        "Тип панели",
        "Маркировка (буква)",
        "Маркировка (номер)",
        "RAL наруж",
        "Толщина металла наруж, мм",
        "Профилирование наруж",
        "RAL внутр",
        "Толщина металла внутр, мм",
        "Профилирование внутр",
        "Длина, мм",
        "Ширина, мм",
        "Толщина, мм",
        "Кол-во, шт.",
        "Площадь, м2 общая",
        "Покрытие наруж",
        "Покрытие внутр",
    ]

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws.cell(row=1, column=1, value=title)
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r, row in enumerate(rows, start=4):
        values = [
            row.idx,
            row.supply_no,
            row.panel_type,
            row.tag_prefix,
            row.tag_number,
            row.ral_out,
            row.metal_out_mm,
            row.profile_out,
            row.ral_in,
            row.metal_in_mm,
            row.profile_in,
            row.length_mm,
            row.width_mm,
            row.thickness_mm,
            row.qty,
            row.area_m2_total,
            row.coating_out,
            row.coating_in,
        ]
        for c, v in enumerate(values, start=1):
            ws.cell(row=r, column=c, value=v)

    total_qty = sum(r.qty for r in rows)
    total_area = sum(r.area_m2_total for r in rows)
    total_row = 4 + len(rows)
    ws.cell(row=total_row, column=1, value="ИТОГО")
    ws.cell(row=total_row, column=1).font = Font(bold=True)
    ws.cell(row=total_row, column=15, value=round(total_qty, 3))
    ws.cell(row=total_row, column=16, value=round(total_area, 3))
    ws.cell(row=total_row, column=15).font = Font(bold=True)
    ws.cell(row=total_row, column=16).font = Font(bold=True)

    ws.freeze_panes = "A4"

    for col in range(1, len(headers) + 1):
        ws.column_dimensions[chr(64 + col)].width = 18
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 20

    wb.save(path)

