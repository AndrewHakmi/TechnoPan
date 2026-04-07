from __future__ import annotations

import os


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

