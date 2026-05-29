# this file is fully vibecoded
from __future__ import annotations

import json
import re
import sys
import threading
from io import StringIO
from pathlib import Path
from tkinter import filedialog as fd

import customtkinter as ctk
from tkinterdnd2 import TkinterDnD, DND_ALL
import pandas as pd

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))
from table_lookup import TableLookup


def _resource(relative: str) -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / relative
    return BASE_DIR / relative

PLACEHOLDER_SHEET = "— лист —"
PLACEHOLDER_COL   = "— столбец —"
AUTO_OUT          = "— авто —"
DATA_EXTS         = {".xls", ".xlsx", ".csv"}
PREFS_PATH        = BASE_DIR / ".gui_prefs.json"


# ── helpers ───────────────────────────────────────────────────────────────────

def _parse_first_path(data: str) -> str:
    data = data.strip()
    m = re.search(r'\{([^}]+)\}', data)
    if m:
        return m.group(1)
    return data.split()[0] if " " in data else data


def _read_meta(path: str | Path) -> tuple[list[str], dict[str, list[str]]]:
    """Returns (sheets, {sheet: [column_names]})."""
    path = Path(path)
    suf = path.suffix.lower()
    if suf == ".csv":
        df = None
        for enc in ("utf-8-sig", "utf-8", "cp1251"):
            try:
                df = pd.read_csv(path, nrows=0, dtype=str, keep_default_na=False, encoding=enc)
                break
            except UnicodeDecodeError:
                pass
        if df is None:
            raise ValueError("Не удалось прочитать CSV")
        return [""], {"": [str(c).strip() for c in df.columns]}
    if suf in (".xlsx", ".xls"):
        xl = pd.ExcelFile(path)
        sheets = xl.sheet_names
        meta = {
            s: [str(c).strip() for c in pd.read_excel(path, sheet_name=s, nrows=0).columns]
            for s in sheets
        }
        return sheets, meta
    raise ValueError(f"Формат {suf!r} не поддерживается")


def _abs(p: str) -> str:
    if not p:
        return p
    ap = Path(p)
    if ap.is_absolute():
        return p
    full = BASE_DIR / p
    return str(full) if full.exists() else p


def _rel(p: str) -> str:
    try:
        return str(Path(p).relative_to(BASE_DIR))
    except ValueError:
        return p


# ── SourcePanel ───────────────────────────────────────────────────────────────

class SourcePanel(ctk.CTkFrame):

    def __init__(self, parent, n: int, **kw):
        super().__init__(parent, **kw)
        self._n = n
        self._path: str | None = None
        self._meta: dict[str, list[str]] = {}
        self._on_delete: callable = lambda: None
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(8, 2))

        self._lbl_n = ctk.CTkLabel(hdr, text=f"Источник {self._n}", font=("Arial", 12, "bold"))
        self._lbl_n.pack(side="left")

        ctk.CTkButton(
            hdr, text="✕", width=26, height=26,
            fg_color="#c0392b", hover_color="#e74c3c",
            command=lambda: self._on_delete(),
        ).pack(side="right")

        fr = ctk.CTkFrame(self, fg_color="transparent")
        fr.pack(fill="x", padx=10, pady=2)

        self._file_lbl = ctk.CTkLabel(
            fr, text="нет файла — перетащите или нажмите 📁",
            text_color="gray", anchor="w",
        )
        self._file_lbl.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            fr, text="📁", width=30, height=26,
            fg_color=("gray70", "gray28"), hover_color=("gray60", "gray38"),
            command=self._pick,
        ).pack(side="left", padx=(6, 0))

        dr = ctk.CTkFrame(self, fg_color="transparent")
        dr.pack(fill="x", padx=10, pady=(3, 2))

        self._sheet_var = ctk.StringVar(value=PLACEHOLDER_SHEET)
        self._sheet_menu = ctk.CTkOptionMenu(
            dr, variable=self._sheet_var, values=[PLACEHOLDER_SHEET],
            width=112, state="disabled", command=self._on_sheet, dynamic_resizing=False,
        )
        self._sheet_menu.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(dr, text="Ключ:", font=("Arial", 11)).pack(side="left", padx=(0, 3))
        self._key_var = ctk.StringVar(value=PLACEHOLDER_COL)
        self._key_menu = ctk.CTkOptionMenu(
            dr, variable=self._key_var, values=[PLACEHOLDER_COL],
            width=128, state="disabled", dynamic_resizing=False,
        )
        self._key_menu.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(dr, text="Знач.:", font=("Arial", 11)).pack(side="left", padx=(0, 3))
        self._val_var = ctk.StringVar(value=PLACEHOLDER_COL)
        self._val_menu = ctk.CTkOptionMenu(
            dr, variable=self._val_var, values=[PLACEHOLDER_COL],
            width=128, state="disabled", dynamic_resizing=False,
        )
        self._val_menu.pack(side="left")

        tr = ctk.CTkFrame(self, fg_color="transparent")
        tr.pack(fill="x", padx=10, pady=(2, 8))

        ctk.CTkLabel(tr, text="Шаблон:", font=("Arial", 11)).pack(side="left", padx=(0, 4))
        self._tmpl_var = ctk.StringVar()
        ctk.CTkEntry(
            tr, textvariable=self._tmpl_var,
            placeholder_text='необязательно, напр.: дата установки: {value}',
            height=28,
        ).pack(side="left", fill="x", expand=True)

    def set_number(self, n: int):
        self._n = n
        self._lbl_n.configure(text=f"Источник {n}")

    def _pick(self):
        p = fd.askopenfilename(filetypes=[("Данные", "*.xlsx *.xls *.csv")])
        if p:
            self.load_file(p)

    def load_file(self, path: str):
        try:
            sheets, meta = _read_meta(path)
        except Exception as e:
            self._file_lbl.configure(text=f"Ошибка: {e}", text_color="#e74c3c")
            return
        self._path = path
        self._meta = meta
        self._file_lbl.configure(text=Path(path).name, text_color=("black", "white"))

        if len(sheets) <= 1:
            self._sheet_menu.configure(values=sheets or [PLACEHOLDER_SHEET], state="disabled")
            self._sheet_var.set(sheets[0] if sheets else PLACEHOLDER_SHEET)
        else:
            self._sheet_menu.configure(values=sheets, state="normal")
            self._sheet_var.set(sheets[0])

        self._on_sheet(self._sheet_var.get())

    def _on_sheet(self, sheet: str):
        cols = self._meta.get(sheet, [])
        if cols:
            self._key_menu.configure(values=cols, state="normal")
            self._key_var.set(cols[0])
            self._val_menu.configure(values=cols, state="normal")
            self._val_var.set(cols[min(1, len(cols) - 1)])
        else:
            for m, v in [(self._key_menu, self._key_var), (self._val_menu, self._val_var)]:
                m.configure(values=[PLACEHOLDER_COL], state="disabled")
                v.set(PLACEHOLDER_COL)

    def get_config(self) -> dict | None:
        if not self._path:
            return None
        key, val = self._key_var.get(), self._val_var.get()
        if PLACEHOLDER_COL in (key, val):
            return None
        sheet = self._sheet_var.get()
        return {
            "path": self._path,
            "sheet": "" if sheet in ("", PLACEHOLDER_SHEET) else sheet,
            "key_column": key,
            "value_column": val,
            "template": self._tmpl_var.get(),
        }

    def apply_config(self, cfg: dict):
        p = _abs(cfg.get("path", ""))
        if not p or not Path(p).exists():
            return
        self.load_file(p)
        s = cfg.get("sheet", "")
        if s and s in self._meta:
            self._sheet_var.set(s)
            self._on_sheet(s)
        cols = self._meta.get(self._sheet_var.get(), [])
        if cfg.get("key_column") in cols:
            self._key_var.set(cfg["key_column"])
        if cfg.get("value_column") in cols:
            self._val_var.set(cfg["value_column"])
        self._tmpl_var.set(cfg.get("template", ""))


# ── MainPanel ─────────────────────────────────────────────────────────────────

class MainPanel(ctk.CTkFrame):

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self._path: str | None = None
        self._meta: dict[str, list[str]] = {}
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text="ГЛАВНАЯ ТАБЛИЦА", font=("Arial", 13, "bold")).pack(
            anchor="w", padx=12, pady=(10, 4)
        )

        fr = ctk.CTkFrame(self, fg_color="transparent")
        fr.pack(fill="x", padx=12, pady=2)

        self._file_lbl = ctk.CTkLabel(
            fr, text="файл не выбран — перетащите или нажмите 📁",
            text_color="gray", anchor="w",
        )
        self._file_lbl.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            fr, text="📁", width=32, height=28,
            fg_color=("gray70", "gray28"), hover_color=("gray60", "gray38"),
            command=self._pick,
        ).pack(side="left", padx=(6, 2))

        ctk.CTkButton(
            fr, text="✕", width=32, height=28,
            fg_color="#c0392b", hover_color="#e74c3c",
            command=self._clear,
        ).pack(side="left")

        dr1 = ctk.CTkFrame(self, fg_color="transparent")
        dr1.pack(fill="x", padx=12, pady=(4, 2))

        ctk.CTkLabel(dr1, text="Лист:", font=("Arial", 11)).pack(side="left", padx=(0, 4))
        self._sheet_var = ctk.StringVar(value=PLACEHOLDER_SHEET)
        self._sheet_menu = ctk.CTkOptionMenu(
            dr1, variable=self._sheet_var, values=[PLACEHOLDER_SHEET],
            width=132, state="disabled", command=self._on_sheet, dynamic_resizing=False,
        )
        self._sheet_menu.pack(side="left", padx=(0, 14))

        ctk.CTkLabel(dr1, text="Ключевой столбец:", font=("Arial", 11)).pack(side="left", padx=(0, 4))
        self._key_var = ctk.StringVar(value=PLACEHOLDER_COL)
        self._key_menu = ctk.CTkOptionMenu(
            dr1, variable=self._key_var, values=[PLACEHOLDER_COL],
            width=156, state="disabled", dynamic_resizing=False,
        )
        self._key_menu.pack(side="left")

        dr2 = ctk.CTkFrame(self, fg_color="transparent")
        dr2.pack(fill="x", padx=12, pady=(2, 10))

        ctk.CTkLabel(dr2, text="Столбец для записи:", font=("Arial", 11)).pack(side="left", padx=(0, 4))
        self._out_var = ctk.StringVar(value=AUTO_OUT)
        self._out_menu = ctk.CTkOptionMenu(
            dr2, variable=self._out_var, values=[AUTO_OUT],
            width=186, state="disabled", dynamic_resizing=False,
        )
        self._out_menu.pack(side="left")

    def _pick(self):
        p = fd.askopenfilename(filetypes=[("Данные", "*.xlsx *.xls *.csv")])
        if p:
            self.load_file(p)

    def _clear(self):
        self._path = None
        self._meta = {}
        self._file_lbl.configure(
            text="файл не выбран — перетащите или нажмите 📁", text_color="gray"
        )
        for m, v, ph in [
            (self._sheet_menu, self._sheet_var, PLACEHOLDER_SHEET),
            (self._key_menu,   self._key_var,   PLACEHOLDER_COL),
            (self._out_menu,   self._out_var,   AUTO_OUT),
        ]:
            m.configure(values=[ph], state="disabled")
            v.set(ph)

    def load_file(self, path: str):
        try:
            sheets, meta = _read_meta(path)
        except Exception as e:
            self._file_lbl.configure(text=f"Ошибка: {e}", text_color="#e74c3c")
            return
        self._path = path
        self._meta = meta
        self._file_lbl.configure(text=Path(path).name, text_color=("black", "white"))

        if len(sheets) <= 1:
            self._sheet_menu.configure(values=sheets or [PLACEHOLDER_SHEET], state="disabled")
            self._sheet_var.set(sheets[0] if sheets else PLACEHOLDER_SHEET)
        else:
            self._sheet_menu.configure(values=sheets, state="normal")
            self._sheet_var.set(sheets[0])

        self._on_sheet(self._sheet_var.get())

    def _on_sheet(self, sheet: str):
        cols = self._meta.get(sheet, [])
        if cols:
            self._key_menu.configure(values=cols, state="normal")
            self._key_var.set(cols[0])
            self._out_menu.configure(values=[AUTO_OUT] + cols, state="normal")
            self._out_var.set(AUTO_OUT)
        else:
            for m, v, ph in [
                (self._key_menu, self._key_var, PLACEHOLDER_COL),
                (self._out_menu, self._out_var, AUTO_OUT),
            ]:
                m.configure(values=[ph], state="disabled")
                v.set(ph)

    def get_config(self) -> dict | None:
        if not self._path:
            return None
        key = self._key_var.get()
        if key == PLACEHOLDER_COL:
            return None
        sheet = self._sheet_var.get()
        out = self._out_var.get()
        return {
            "path": self._path,
            "sheet": "" if sheet in ("", PLACEHOLDER_SHEET) else sheet,
            "key_column": key,
            "output_column": "" if out == AUTO_OUT else out,
        }

    def apply_config(self, cfg: dict):
        p = _abs(cfg.get("path", ""))
        if not p or not Path(p).exists():
            return
        self.load_file(p)
        s = cfg.get("sheet", "")
        if s and s in self._meta:
            self._sheet_var.set(s)
            self._on_sheet(s)
        cols = self._meta.get(self._sheet_var.get(), [])
        if cfg.get("key_column") in cols:
            self._key_var.set(cfg["key_column"])
        oc = cfg.get("output_column", "")
        if oc and oc in cols:
            self._out_var.set(oc)


# ── App ───────────────────────────────────────────────────────────────────────

class App(ctk.CTk, TkinterDnD.DnDWrapper):

    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)

        self._sources: list[SourcePanel] = []
        self._running = False

        prefs = self._load_prefs()
        mode = prefs.get("appearance_mode", "dark")
        ctk.set_appearance_mode(mode)
        ctk.set_default_color_theme("blue")

        self.title("SimpleExcelCopypaster")
        self.geometry("640x720")
        self.minsize(580, 500)

        self._build_ui()
        self._setup_dnd()

    # ── prefs ─────────────────────────────────────────────────────────────────

    def _load_prefs(self) -> dict:
        if PREFS_PATH.exists():
            try:
                return json.loads(PREFS_PATH.read_text())
            except Exception:
                pass
        return {}

    def _save_prefs(self, **kw):
        prefs = self._load_prefs()
        prefs.update(kw)
        try:
            PREFS_PATH.write_text(json.dumps(prefs))
        except Exception:
            pass

    # ── UI construction ───────────────────────────────────────────────────────

    def _build_ui(self):
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=(12, 0))

        logo_path = _resource("resources/logo.png")
        if logo_path.exists():
            from PIL import Image
            self._logo_img = ctk.CTkImage(Image.open(logo_path), size=(36, 36))
            ctk.CTkLabel(hdr, image=self._logo_img, text="").pack(side="left", padx=(0, 8))

        ctk.CTkLabel(hdr, text="SimpleExcelCopypaster", font=("Arial", 15, "bold")).pack(side="left")

        mode = ctk.get_appearance_mode().lower()
        self._theme_btn = ctk.CTkButton(
            hdr, text="☀" if mode == "dark" else "🌙",
            width=40, height=36, font=("Arial", 16),
            fg_color=("gray70", "gray28"), hover_color=("gray60", "gray38"),
            command=self._toggle_theme,
        )
        self._theme_btn.pack(side="right")

        self._main = MainPanel(self, border_width=1, corner_radius=8)
        self._main.pack(fill="x", padx=20, pady=(12, 0))

        src_hdr = ctk.CTkFrame(self, fg_color="transparent")
        src_hdr.pack(fill="x", padx=20, pady=(14, 4))
        ctk.CTkLabel(src_hdr, text="ИСТОЧНИКИ", font=("Arial", 13, "bold")).pack(side="left")
        ctk.CTkButton(
            src_hdr, text="+ Добавить источник", height=30,
            fg_color=("gray70", "gray28"), hover_color=("gray60", "gray38"),
            command=lambda: self._add_source(),
        ).pack(side="right")

        self._scroll = ctk.CTkScrollableFrame(self, corner_radius=8, border_width=1)
        self._scroll.pack(fill="both", expand=True, padx=20, pady=0)

        self._hint = ctk.CTkLabel(
            self._scroll,
            text="Перетащите файл сюда или нажмите «+ Добавить источник»",
            text_color="gray", font=("Arial", 12),
        )
        self._hint.pack(pady=28)

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(10, 0))

        self._run_btn = ctk.CTkButton(
            btn_row, text="▶  ЗАПУСТИТЬ", height=44, font=("Arial", 14, "bold"),
            fg_color="#27ae60", hover_color="#2ecc71",
            command=self._run,
        )
        self._run_btn.pack(fill="x")

        self._status = ctk.CTkLabel(self, text="", font=("Arial", 12), wraplength=580)
        self._status.pack(pady=(6, 12))

    def _setup_dnd(self):
        self._overlay = ctk.CTkFrame(
            self, fg_color=("gray95", "gray5"),
            border_width=4, border_color="#3498db", corner_radius=16,
        )
        ctk.CTkLabel(
            self._overlay,
            text="Бросьте файл данных сюда\n(Excel или CSV)",
            font=("Arial", 20, "bold"), text_color="#3498db",
        ).place(relx=0.5, rely=0.5, anchor="center")

        self.drop_target_register(DND_ALL)
        self.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.dnd_bind("<<DropLeave>>", lambda _e: self._overlay.place_forget())
        self.dnd_bind("<<Drop>>", self._on_drop)

    # ── theme ─────────────────────────────────────────────────────────────────

    def _toggle_theme(self):
        new = "light" if ctk.get_appearance_mode() == "Dark" else "dark"
        ctk.set_appearance_mode(new)
        self._theme_btn.configure(text="☀" if new == "dark" else "🌙")
        self._save_prefs(appearance_mode=new)

    # ── sources management ────────────────────────────────────────────────────

    def _add_source(self, path: str | None = None) -> SourcePanel:
        self._hint.pack_forget()
        n = len(self._sources) + 1
        panel = SourcePanel(
            self._scroll, n=n,
            fg_color=("gray85", "gray22"),
            corner_radius=6, border_width=1,
        )
        panel._on_delete = lambda: self._remove_source(panel)
        panel.pack(fill="x", pady=(0, 6), padx=2)
        self._sources.append(panel)
        if path:
            panel.load_file(path)
        return panel

    def _remove_source(self, panel: SourcePanel):
        panel.pack_forget()
        panel.destroy()
        self._sources.remove(panel)
        for i, p in enumerate(self._sources, 1):
            p.set_number(i)
        if not self._sources:
            self._hint.pack(pady=28)

    # ── drag-n-drop ───────────────────────────────────────────────────────────

    def _on_drag_enter(self, _event):
        self._overlay.place(relx=0.02, rely=0.02, relwidth=0.96, relheight=0.96)
        self._overlay.lift()
        self.update_idletasks()

    def _on_drop(self, event):
        self._overlay.place_forget()
        self.focus_force()

        path = _parse_first_path(event.data)
        if Path(path).suffix.lower() not in DATA_EXTS:
            self._set_status("Неверный формат — ожидается .xlsx, .xls или .csv", "red")
            return

        if not self._main._path:
            self._main.load_file(path)
        else:
            self._add_source(path)

    # ── config ────────────────────────────────────────────────────────────────

    def _build_config(self) -> dict | None:
        main = self._main.get_config()
        if main is None:
            self._set_status("Главная таблица не настроена", "red")
            return None
        if not self._sources:
            self._set_status("Добавьте хотя бы один источник", "red")
            return None
        srcs = []
        for i, panel in enumerate(self._sources, 1):
            c = panel.get_config()
            if c is None:
                self._set_status(f"Источник {i}: файл или столбцы не выбраны", "red")
                return None
            srcs.append(c)
        return {"main": main, "sources": srcs}

    # ── run ───────────────────────────────────────────────────────────────────

    def _run(self):
        if self._running:
            return
        cfg = self._build_config()
        if cfg is None:
            return
        self._running = True
        self._run_btn.configure(text="⏳  Работаю...", state="disabled")
        threading.Thread(target=self._execute, args=(cfg,), daemon=True).start()

    def _execute(self, cfg: dict):
        buf = StringIO()
        old_stdout = sys.stdout
        sys.stdout = buf
        try:
            m = cfg["main"]
            lookup = TableLookup(
                path=m["path"],
                key_column=m["key_column"],
                sheet_name=m["sheet"] or None,
                output_column=m["output_column"] or None,
            )
            for src in cfg["sources"]:
                lookup.addSourceTable(
                    path=src["path"],
                    key_column=src["key_column"],
                    value_column=src["value_column"],
                    sheet_name=src["sheet"] or None,
                    template=src["template"] or None,
                )
            lookup.run()
            output = buf.getvalue().strip()
            msg = output.split("\n")[-1] if output else "Готово!"
            self.after(0, lambda: self._set_status(msg, "green"))
        except Exception as e:
            err = str(e)
            self.after(0, lambda: self._set_status(f"Ошибка: {err}", "red"))
        finally:
            sys.stdout = old_stdout
            self._running = False
            self.after(0, lambda: self._run_btn.configure(text="▶  ЗАПУСТИТЬ", state="normal"))

    # ── status ────────────────────────────────────────────────────────────────

    def _set_status(self, text: str, kind: str = "gray"):
        colors = {
            "red":    "#e74c3c",
            "green":  "#2ecc71",
            "orange": "#f39c12",
            "gray":   ("gray60", "gray40"),
        }
        self._status.configure(text=text, text_color=colors.get(kind, ("black", "white")))


if __name__ == "__main__":
    app = App()
    app.mainloop()
