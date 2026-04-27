from __future__ import annotations

import os
import subprocess
import tempfile
import shutil
from pathlib import Path


def resolve_odafc_win_exec_path() -> str | None:
    env_path = os.environ.get("ODA_FILE_CONVERTER_EXE")
    if env_path:
        return env_path

    pf = os.environ.get("ProgramFiles")
    pfx86 = os.environ.get("ProgramFiles(x86)")
    bases = [b for b in [pf, pfx86] if b]

    direct_candidates: list[str] = []
    oda_roots: list[str] = []
    for base in bases:
        direct_candidates.append(os.path.join(base, "ODA", "ODAFileConverter", "ODAFileConverter.exe"))
        direct_candidates.append(os.path.join(base, "ODA", "ODAFile Converter", "ODAFileConverter.exe"))
        oda_roots.append(os.path.join(base, "ODA"))

    for p in direct_candidates:
        if os.path.exists(p):
            return p

    for root in oda_roots:
        if not os.path.isdir(root):
            continue
        try:
            subdirs = [
                os.path.join(root, name)
                for name in os.listdir(root)
                if os.path.isdir(os.path.join(root, name))
            ]
        except OSError:
            continue
        for d in subdirs:
            if "odafileconverter" not in os.path.basename(d).lower():
                continue
            exe = os.path.join(d, "ODAFileConverter.exe")
            if os.path.exists(exe):
                return exe

    return None


def safe_convert_dwg_to_dxf(dwg_path: Path, out_version: str = "ACAD2018") -> Path:
    """
    Безопасная конвертация DWG в DXF через ODA File Converter.
    Избегает зависания (deadlock), которое есть во встроенном модуле ezdxf.addons.odafc
    при запуске из PyInstaller windowed mode.
    """
    exe = resolve_odafc_win_exec_path()
    if not exe:
        raise RuntimeError("ODA File Converter не найден.")

    # ODAFC требует абсолютных путей к директориям
    dwg_abs = dwg_path.resolve()
    in_folder = str(dwg_abs.parent)
    filename = dwg_abs.name

    tmp_dir = tempfile.mkdtemp(prefix="technopan_odafc_")
    
    # Аргументы ODAFC: InputFolder OutputFolder Version Format Recurse Audit [Filename]
    cmd = [
        exe,
        in_folder,
        tmp_dir,
        out_version,
        "DXF",
        "0",  # Recurse
        "1",  # Audit
        filename
    ]

    try:
        # Используем CREATE_NO_WINDOW чтобы скрыть консоль на Windows
        # capture_output=True автоматически читает stdout/stderr и предотвращает deadlock
        subprocess.run(
            cmd,
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0,
            check=True
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Ошибка при конвертации ODAFC: {e.stderr.decode('utf-8', errors='ignore')}")
    except Exception as e:
        raise RuntimeError(f"Не удалось запустить ODAFC: {e}")

    # Ищем результат в tmp_dir
    dxf_name = dwg_abs.with_suffix(".dxf").name
    result_path = Path(tmp_dir) / dxf_name

    if not result_path.exists():
        # Иногда ODAFC меняет регистр расширения
        found = list(Path(tmp_dir).glob("*.[dD][xX][fF]"))
        if found:
            result_path = found[0]
        else:
            raise RuntimeError("Конвертация прошла без ошибок, но DXF файл не был создан.")

    return result_path

