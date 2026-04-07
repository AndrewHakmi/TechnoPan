import argparse
import json
from pathlib import Path

from .config import load_config
from .dxf import inspect_dxf, extract_panels_from_dxf
from .odafc_utils import resolve_odafc_win_exec_path
from .spec import build_panel_rows, write_spec_xlsx


def main() -> None:
    parser = argparse.ArgumentParser(prog="technopan_spec")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_inspect = sub.add_parser("inspect", help="Показать блоки и атрибуты в DXF")
    p_inspect.add_argument("path", type=Path)
    p_inspect.add_argument("--json", action="store_true")

    p_gen = sub.add_parser("generate", help="Сгенерировать спецификацию панелей в .xlsx")
    p_gen.add_argument("path", type=Path)
    p_gen.add_argument("-c", "--config", type=Path, default=Path("configs/default.yml"))
    p_gen.add_argument("-o", "--out", type=Path, default=Path("spec.xlsx"))
    p_gen.add_argument("--title", type=str, default="Коммерческое предложение")

    p_odafc = sub.add_parser("odafc-check", help="Проверить доступность ODA File Converter")
    p_odafc.add_argument("--print-setx", action="store_true")

    sub.add_parser("gui", help="Запустить графический интерфейс")

    args = parser.parse_args()

    if args.cmd == "gui":
        from .gui import App
        app = App()
        app.mainloop()
        return

    if args.cmd == "inspect":
        result = inspect_dxf(args.path)
        if args.json:
            print(json.dumps(result, ensure_ascii=False, indent=2))
        else:
            for block_name, data in sorted(result.items()):
                print(block_name)
                print(f"  inserts: {data['count']}")
                if data["attr_tags"]:
                    print("  attr_tags:")
                    for t in sorted(data["attr_tags"]):
                        print(f"    - {t}")
                else:
                    print("  attr_tags: (none)")
        return

    if args.cmd == "generate":
        cfg = load_config(args.config)
        panels = extract_panels_from_dxf(args.path, cfg)
        rows = build_panel_rows(panels)
        write_spec_xlsx(args.out, rows, title=args.title)
        print(str(args.out))
        return

    if args.cmd == "odafc-check":
        import ezdxf
        from ezdxf.addons import odafc

        candidate = resolve_odafc_win_exec_path()
        if candidate:
            ezdxf.options.set("odafc-addon", "win_exec_path", candidate)

        installed = odafc.is_installed()
        print(f"installed: {installed}")
        print(f"ODA_FILE_CONVERTER_EXE: {candidate or ''}")
        if args.print_setx and candidate:
            print(f"setx ODA_FILE_CONVERTER_EXE \"{candidate}\"")
        return

    raise SystemExit(2)

