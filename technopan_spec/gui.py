
from __future__ import annotations

import sys
import threading
import time
import traceback
from pathlib import Path
from tkinter import BooleanVar, filedialog, messagebox

import customtkinter as ctk

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ---------------------------------------------------------------------------
# Resolve base directory so paths work both from source and PyInstaller .exe
# ---------------------------------------------------------------------------
if getattr(sys, "frozen", False):
    # Running as PyInstaller bundle
    _BASE_DIR = Path(sys._MEIPASS)  # type: ignore[attr-defined]
else:
    _BASE_DIR = Path(__file__).parent.parent  # project root

_CONFIGS_DIR = _BASE_DIR / "configs"

# ---------------------------------------------------------------------------
# Table column definitions:  (PanelRow attr, header, min-width px)
# ---------------------------------------------------------------------------
_COLUMNS: list[tuple[str, str, int]] = [
    ("panel_type",    "Тип панели",    140),
    ("tag_prefix",    "Букв.",          50),
    ("tag_number",    "Номер",          60),
    ("length_mm",     "Длина, мм",      80),
    ("width_mm",      "Ширина, мм",     80),
    ("thickness_mm",  "Толщина, мм",    80),
    ("ral_out",       "RAL нар.",        70),
    ("ral_in",        "RAL вн.",         70),
    ("qty",           "Кол-во",          60),
    ("area_m2_total", "Площадь, м²",     90),
]


class _RowWidget:
    """One data row inside the scrollable review table."""

    def __init__(self, parent: ctk.CTkScrollableFrame, row_idx: int, panel_row, on_toggle):
        self.row = panel_row
        self.var = BooleanVar(value=True)

        bg = ("gray90", "gray20") if row_idx % 2 == 0 else ("gray85", "gray17")
        self.frame = ctk.CTkFrame(parent, fg_color=bg, corner_radius=0)
        self.frame.grid(row=row_idx, column=0, sticky="ew")
        self.frame.grid_columnconfigure(0, minsize=34)
        for ci, (_, _, w) in enumerate(_COLUMNS, start=1):
            self.frame.grid_columnconfigure(ci, minsize=w)

        self.cb = ctk.CTkCheckBox(
            self.frame, text="", variable=self.var,
            width=24, height=24, command=on_toggle,
        )
        self.cb.grid(row=0, column=0, padx=(6, 2), pady=3)

        values = [
            panel_row.panel_type,
            panel_row.tag_prefix or "—",
            panel_row.tag_number if panel_row.tag_number is not None else "—",
            int(panel_row.length_mm),
            int(panel_row.width_mm),
            int(panel_row.thickness_mm),
            panel_row.ral_out or "—",
            panel_row.ral_in or "—",
            panel_row.qty,
            panel_row.area_m2_total,
        ]
        for ci, (val, (_, _, w)) in enumerate(zip(values, _COLUMNS), start=1):
            ctk.CTkLabel(
                self.frame, text=str(val),
                width=w - 4, anchor="center",
            ).grid(row=0, column=ci, padx=2, pady=3)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("TechnoPan Specification Generator")
        self.geometry("1150x860")
        self.minsize(900, 640)

        self._all_rows: list = []
        self._row_widgets: list[_RowWidget] = []
        self._stop_event: threading.Event | None = None
        self._extract_start_time: float = 0.0

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()

    # ─────────────────────────────────────── sidebar
    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=180, corner_radius=0)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            sb, text="TechnoPan\nGenerator",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, padx=16, pady=(24, 12))

        ctk.CTkLabel(sb, text="Внешний вид:", anchor="w").grid(
            row=5, column=0, padx=16, pady=(8, 0), sticky="w"
        )
        menu = ctk.CTkOptionMenu(
            sb, values=["System", "Light", "Dark"],
            command=lambda v: ctk.set_appearance_mode(v),
        )
        menu.set("System")
        menu.grid(row=6, column=0, padx=16, pady=(4, 16), sticky="ew")

    # ─────────────────────────────────────── main layout
    def _build_main(self):
        mf = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        mf.grid(row=0, column=1, sticky="nsew", padx=12, pady=12)
        mf.grid_columnconfigure(0, weight=1)
        self._main = mf

        # ── Step 1 ──────────────────────────────────────────────────────────
        ctk.CTkLabel(
            mf, text="Шаг 1 — Входные данные",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=4, pady=(4, 6))

        # File row
        ff = ctk.CTkFrame(mf, fg_color="transparent")
        ff.grid(row=1, column=0, sticky="ew", padx=4, pady=4)
        ff.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(ff, text="Выбрать файл", width=130, command=self._select_file).grid(
            row=0, column=0, padx=(0, 8)
        )
        self._entry_file = ctk.CTkEntry(ff, placeholder_text="Файл не выбран")
        self._entry_file.grid(row=0, column=1, sticky="ew")

        # Config + auto-detect row
        cf = ctk.CTkFrame(mf, fg_color="transparent")
        cf.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
        cf.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(cf, text="Конфигурация:", width=130, anchor="w").grid(
            row=0, column=0, padx=(0, 8)
        )
        configs = self._get_config_files()
        self._option_config = ctk.CTkOptionMenu(cf, values=configs)
        if "text_tags.yml" in configs:
            self._option_config.set("text_tags.yml")
        elif configs:
            self._option_config.set(configs[0])
        self._option_config.grid(row=0, column=1, sticky="ew", padx=(0, 8))

        self._btn_autodetect = ctk.CTkButton(
            cf, text="Определить автоматически",
            width=190,
            fg_color="#555", hover_color="#333",
            command=self._start_autodetect,
        )
        self._btn_autodetect.grid(row=0, column=2)

        # Extract + Stop row
        action_row = ctk.CTkFrame(mf, fg_color="transparent")
        action_row.grid(row=3, column=0, sticky="ew", padx=4, pady=(10, 4))
        action_row.grid_columnconfigure(0, weight=1)

        self._btn_extract = ctk.CTkButton(
            action_row,
            text="ИЗВЛЕЧЬ ПАНЕЛИ",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44, fg_color="#1a6fb5", hover_color="#145a94",
            command=self._start_extract,
        )
        self._btn_extract.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self._btn_stop = ctk.CTkButton(
            action_row,
            text="⏹ Остановить",
            width=130, height=44,
            fg_color="#8b0000", hover_color="#660000",
            state="disabled",
            command=self._stop_extraction,
        )
        self._btn_stop.grid(row=0, column=1)

        # ── Step 2 ──────────────────────────────────────────────────────────
        ctk.CTkLabel(
            mf, text="Шаг 2 — Найденные панели",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=4, column=0, sticky="w", padx=4, pady=(16, 2))

        sel_row = ctk.CTkFrame(mf, fg_color="transparent")
        sel_row.grid(row=5, column=0, sticky="ew", padx=4, pady=(0, 4))
        ctk.CTkButton(sel_row, text="Выбрать все",  width=120, command=self._select_all).grid(row=0, column=0, padx=(0, 6))
        ctk.CTkButton(sel_row, text="Снять все",    width=120, command=self._deselect_all).grid(row=0, column=1, padx=(0, 16))
        self._lbl_summary = ctk.CTkLabel(sel_row, text="нет данных", anchor="w")
        self._lbl_summary.grid(row=0, column=2, sticky="w")

        # Table header
        hdr = ctk.CTkFrame(mf, corner_radius=4)
        hdr.grid(row=6, column=0, sticky="ew", padx=4)
        hdr.grid_columnconfigure(0, minsize=34)
        for ci, (_, label, w) in enumerate(_COLUMNS, start=1):
            hdr.grid_columnconfigure(ci, minsize=w)
            ctk.CTkLabel(
                hdr, text=label, width=w - 4,
                font=ctk.CTkFont(weight="bold"), anchor="center",
            ).grid(row=0, column=ci, padx=2, pady=6)

        # Table body
        self._tbl_body = ctk.CTkScrollableFrame(mf, height=260, corner_radius=4)
        self._tbl_body.grid(row=7, column=0, sticky="ew", padx=4)
        self._tbl_body.grid_columnconfigure(0, weight=1)

        self._lbl_empty = ctk.CTkLabel(
            self._tbl_body,
            text="Сначала извлеките панели (Шаг 1)",
            text_color="gray",
        )
        self._lbl_empty.grid(row=0, column=0, pady=20)

        # ── Step 3 ──────────────────────────────────────────────────────────
        ctk.CTkLabel(
            mf, text="Шаг 3 — Параметры вывода",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=8, column=0, sticky="w", padx=4, pady=(16, 6))

        of = ctk.CTkFrame(mf, fg_color="transparent")
        of.grid(row=9, column=0, sticky="ew", padx=4, pady=4)
        of.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(of, text="Папка вывода", width=130, command=self._select_output_folder).grid(row=0, column=0, padx=(0, 8))
        self._entry_out_folder = ctk.CTkEntry(of, placeholder_text="Папка файла-источника")
        self._entry_out_folder.grid(row=0, column=1, sticky="ew")

        nf = ctk.CTkFrame(mf, fg_color="transparent")
        nf.grid(row=10, column=0, sticky="ew", padx=4, pady=4)
        nf.grid_columnconfigure(1, weight=1)
        nf.grid_columnconfigure(3, weight=2)
        ctk.CTkLabel(nf, text="Имя файла:", width=130, anchor="w").grid(row=0, column=0, padx=(0, 8))
        self._entry_filename = ctk.CTkEntry(nf, placeholder_text="spec.xlsx")
        self._entry_filename.insert(0, "spec.xlsx")
        self._entry_filename.grid(row=0, column=1, sticky="ew", padx=(0, 16))
        ctk.CTkLabel(nf, text="Заголовок:", width=80, anchor="w").grid(row=0, column=2, padx=(0, 8))
        self._entry_title = ctk.CTkEntry(nf, placeholder_text="Коммерческое предложение")
        self._entry_title.insert(0, "Коммерческое предложение")
        self._entry_title.grid(row=0, column=3, sticky="ew")

        self._btn_generate = ctk.CTkButton(
            mf,
            text="СГЕНЕРИРОВАТЬ EXCEL",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=44, fg_color="green", hover_color="darkgreen",
            state="disabled",
            command=self._start_generate,
        )
        self._btn_generate.grid(row=11, column=0, sticky="ew", padx=4, pady=(10, 4))

        # ── Log console ──────────────────────────────────────────────────────
        log_hdr = ctk.CTkFrame(mf, fg_color="transparent")
        log_hdr.grid(row=12, column=0, sticky="ew", padx=4, pady=(12, 2))
        ctk.CTkLabel(log_hdr, text="Журнал:", anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            log_hdr, text="Очистить", width=80, height=24,
            command=self._clear_log,
        ).grid(row=0, column=1, padx=(8, 0))

        self._console = ctk.CTkTextbox(mf, height=180, state="disabled")
        self._console.grid(row=13, column=0, sticky="ew", padx=4, pady=(0, 8))

    # ─────────────────────────────────────── helpers

    def _get_config_files(self) -> list[str]:
        if not _CONFIGS_DIR.exists():
            return ["default.yml"]
        files = sorted(f.name for f in _CONFIGS_DIR.glob("*.yml"))
        return files or ["default.yml"]

    def _elapsed(self) -> str:
        s = int(time.time() - self._extract_start_time)
        return f"[{s//60:02d}:{s%60:02d}]"

    def _log(self, msg: str):
        elapsed = self._elapsed() if self._extract_start_time else ""
        prefix = f"{elapsed} " if elapsed else ""

        def _do():
            self._console.configure(state="normal")
            for line in msg.splitlines():
                self._console.insert("end", prefix + line + "\n")
            self._console.see("end")
            self._console.configure(state="disabled")
        self.after(0, _do)

    def _clear_log(self):
        self._console.configure(state="normal")
        self._console.delete("1.0", "end")
        self._console.configure(state="disabled")

    def _select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("CAD файлы", "*.dwg *.dxf *.dxb"),
                ("DWG файлы", "*.dwg"),
                ("DXF файлы", "*.dxf"),
                ("Все файлы", "*.*"),
            ]
        )
        if path:
            self._entry_file.delete(0, "end")
            self._entry_file.insert(0, path)
            stem = Path(path).stem
            self._entry_filename.delete(0, "end")
            self._entry_filename.insert(0, f"{stem}_spec.xlsx")

    def _select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self._entry_out_folder.delete(0, "end")
            self._entry_out_folder.insert(0, folder)

    # ─────────────────────────────────────── auto-detect

    def _start_autodetect(self):
        path_str = self._entry_file.get().strip()
        if not path_str:
            messagebox.showwarning("Нет файла", "Сначала выберите файл.")
            return
        if not Path(path_str).exists():
            messagebox.showerror("Файл не найден", path_str)
            return
        self._btn_autodetect.configure(state="disabled", text="Определяю…")
        self._extract_start_time = time.time()
        self._stop_event = threading.Event()
        self._btn_stop.configure(state="normal")
        threading.Thread(target=self._autodetect, daemon=True).start()

    def _autodetect(self):
        path_str = self._entry_file.get().strip()
        self._log("═" * 50)
        self._log(f"Автоопределение: {Path(path_str).name}")
        try:
            from technopan_spec.dxf import auto_detect_config
            result = auto_detect_config(
                Path(path_str),
                progress_cb=self._log,
                stop_event=self._stop_event,
            )
            self._log("─" * 40)
            for line in result.details:
                self._log(line)
            self._log("─" * 40)
            self._log(f"✓ Рекомендован конфиг: {result.recommended_config}")

            # Switch config dropdown if recommended file exists
            configs = self._get_config_files()
            if result.recommended_config in configs:
                self.after(0, lambda r=result.recommended_config: self._option_config.set(r))
            else:
                self._log(f"⚠ Файл {result.recommended_config} не найден в папке configs/")

        except InterruptedError:
            self._log("⏹ Остановлено пользователем.")
        except Exception as exc:
            self._log(f"ОШИБКА при автоопределении: {exc}")
            self._log(traceback.format_exc())
        finally:
            self._stop_event = None
            self._extract_start_time = 0.0
            self.after(0, lambda: self._btn_autodetect.configure(
                state="normal", text="Определить автоматически"
            ))
            self.after(0, lambda: self._btn_stop.configure(state="disabled"))

    # ─────────────────────────────────────── table

    def _clear_table(self):
        for w in self._row_widgets:
            w.frame.destroy()
        self._row_widgets.clear()
        self._all_rows.clear()

    def _populate_table(self, rows: list):
        self.after(0, lambda: self._do_populate(rows))

    def _do_populate(self, rows: list):
        self._clear_table()
        self._lbl_empty.grid_remove()

        if not rows:
            self._lbl_empty.configure(
                text="Панели не найдены. Смотрите журнал для диагностики."
            )
            self._lbl_empty.grid(row=0, column=0, pady=20)
            self._update_summary()
            return

        self._all_rows = rows
        for i, row in enumerate(rows):
            rw = _RowWidget(self._tbl_body, i, row, self._update_summary)
            rw.frame.grid_columnconfigure(0, weight=1)
            self._row_widgets.append(rw)
        self._update_summary()

    def _select_all(self):
        for rw in self._row_widgets:
            rw.var.set(True)
        self._update_summary()

    def _deselect_all(self):
        for rw in self._row_widgets:
            rw.var.set(False)
        self._update_summary()

    def _update_summary(self):
        selected = [rw for rw in self._row_widgets if rw.var.get()]
        total = len(self._row_widgets)
        qty = sum(rw.row.qty for rw in selected)
        area = sum(rw.row.area_m2_total for rw in selected)
        self._lbl_summary.configure(
            text=(
                f"Выбрано: {len(selected)} из {total}  │  "
                f"кол-во: {round(qty,1)} шт.  │  "
                f"площадь: {round(area,2)} м²"
            )
        )
        self._btn_generate.configure(
            state="normal" if selected else "disabled"
        )

    def _get_selected_rows(self) -> list:
        return [rw.row for rw in self._row_widgets if rw.var.get()]

    # ─────────────────────────────────────── extract

    def _stop_extraction(self):
        if self._stop_event:
            self._stop_event.set()
            self._btn_stop.configure(state="disabled")
            self._log("⏹ Отправлен сигнал остановки…")

    def _start_extract(self):
        path_str = self._entry_file.get().strip()
        if not path_str:
            messagebox.showwarning("Нет файла", "Сначала выберите файл (Шаг 1).")
            return
        if not Path(path_str).exists():
            messagebox.showerror("Файл не найден", path_str)
            return

        cfg_name = self._option_config.get()
        cfg_path = _CONFIGS_DIR / cfg_name
        if not cfg_path.exists():
            messagebox.showerror(
                "Конфиг не найден",
                f"Файл конфигурации не найден:\n{cfg_path}\n\n"
                f"Ожидаемая папка: {_CONFIGS_DIR}"
            )
            return

        self._extract_start_time = time.time()
        self._stop_event = threading.Event()
        self._btn_extract.configure(state="disabled", text="Извлечение…")
        self._btn_stop.configure(state="normal")
        self._btn_generate.configure(state="disabled")
        threading.Thread(target=self._extract, daemon=True).start()

    def _extract(self):
        path_str = self._entry_file.get().strip()
        config_name = self._option_config.get()
        config_path = _CONFIGS_DIR / config_name

        self._log("═" * 50)
        self._log(f"Файл:    {Path(path_str).name}")
        self._log(f"Конфиг:  {config_name}  ({config_path})")

        try:
            from technopan_spec.config import load_config
            from technopan_spec.dxf import extract_panels_from_dxf
            from technopan_spec.spec import build_panel_rows

            self._log("Загрузка конфигурации…")
            cfg = load_config(config_path)

            mode = (
                "tag" if cfg.tag_extraction.enabled
                else "dimension" if cfg.dimension_extraction.enabled
                else "attribute"
            )
            self._log(f"Режим извлечения: {mode}")

            self._log("Извлечение панелей…")
            panels = extract_panels_from_dxf(
                Path(path_str), cfg,
                progress_cb=self._log,
                stop_event=self._stop_event,
            )
            self._log(f"Найдено объектов: {len(panels)}")

            if not panels:
                self._populate_table([])
                return

            self._log("Группировка…")
            rows = build_panel_rows(panels)
            self._log(f"Уникальных строк спецификации: {len(rows)}")

            self._populate_table(rows)
            self._log("✓ Готово. Выберите строки и нажмите «Сгенерировать Excel».")

        except InterruptedError:
            self._log("⏹ Остановлено пользователем.")
            self._populate_table([])
        except Exception as exc:
            self._log(f"ОШИБКА: {exc}")
            self._log(traceback.format_exc())
            self.after(0, lambda: messagebox.showerror("Ошибка извлечения", str(exc)))
            self._populate_table([])
        finally:
            self._stop_event = None
            self._extract_start_time = 0.0
            self.after(0, lambda: self._btn_extract.configure(
                state="normal", text="ИЗВЛЕЧЬ ПАНЕЛИ"
            ))
            self.after(0, lambda: self._btn_stop.configure(state="disabled"))

    # ─────────────────────────────────────── generate

    def _start_generate(self):
        selected = self._get_selected_rows()
        if not selected:
            messagebox.showwarning("Нет выбранных строк", "Выберите хотя бы одну строку.")
            return
        self._btn_generate.configure(state="disabled", text="Запись…")
        threading.Thread(target=self._generate, args=(selected,), daemon=True).start()

    def _generate(self, rows: list):
        src_path = Path(self._entry_file.get().strip())
        out_folder_str = self._entry_out_folder.get().strip()
        out_folder = Path(out_folder_str) if out_folder_str else src_path.parent
        filename = self._entry_filename.get().strip() or "spec.xlsx"
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
        out_path = out_folder / filename
        title = self._entry_title.get().strip() or "Коммерческое предложение"

        self._log("─" * 40)
        self._log(f"Экспорт: {len(rows)} строк → {out_path}")
        try:
            from technopan_spec.spec import write_spec_xlsx
            write_spec_xlsx(out_path, rows, title=title)
            self._log(f"✓ Файл сохранён: {out_path}")
            self.after(0, lambda: messagebox.showinfo("Готово", f"Файл сохранён:\n{out_path}"))
        except Exception as exc:
            self._log(f"ОШИБКА при записи: {exc}")
            self._log(traceback.format_exc())
            self.after(0, lambda: messagebox.showerror("Ошибка", str(exc)))
        finally:
            can = bool(self._get_selected_rows())
            self.after(0, lambda: self._btn_generate.configure(
                state="normal" if can else "disabled",
                text="СГЕНЕРИРОВАТЬ EXCEL",
            ))


if __name__ == "__main__":
    app = App()
    app.mainloop()
