import json
import sys
from pathlib import Path
from table_lookup import TableLookup

# при exe путь берём от исполняемого файла
# при обычном запуске — от main.py
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).parent

CONFIG_PATH = BASE_DIR / "config.json"

EMPTY_CONFIG = {
    "main": {
        "path": "",
        "sheet": "",
        "key_column": "",
        "output_column": ""
    },
    "sources": [
        {
            "path": "",
            "sheet": "",
            "key_column": "",
            "value_column": "",
            "template": ""
        }
    ]
}


def load_config():
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(
            json.dumps(EMPTY_CONFIG, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"пусто: config.json не найден — создан шаблон: {CONFIG_PATH}")
        return None

    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise ValueError(f"config.json содержит невалидный JSON: {e}")

    main_path = cfg.get("main", {}).get("path", "")
    if not main_path:
        print(f"пусто: config.json есть, но поля не заполнены: {CONFIG_PATH}")
        return None

    return cfg


def _require(obj, field, location):
    val = obj.get(field, "")
    if not val:
        raise ValueError(f"{location}.{field}: обязательное поле пустое или отсутствует")
    return val


def run(cfg: dict | None = None):
    """Запустить алгоритм. cfg — словарь с ключами main/sources; если None — читает config.json."""
    if cfg is None:
        cfg = load_config()
    if cfg is None:
        return

    m = cfg["main"]
    _require(m, "path", "main")
    _require(m, "key_column", "main")

    sources = cfg.get("sources", [])
    if not sources:
        raise ValueError("sources: список пустой, нужен хотя бы один источник")
    for i, src in enumerate(sources):
        _require(src, "path",         f"sources[{i}]")
        _require(src, "key_column",   f"sources[{i}]")
        _require(src, "value_column", f"sources[{i}]")

    lookup = TableLookup(
        path=BASE_DIR / m["path"],
        key_column=m["key_column"],
        sheet_name=m.get("sheet") or None,
        output_column=m.get("output_column") or None,
    )

    for src in sources:
        lookup.addSourceTable(
            path=BASE_DIR / src["path"],
            key_column=src["key_column"],
            value_column=src["value_column"],
            sheet_name=src.get("sheet") or None,
            template=src.get("template") or None,
        )

    lookup.run()


if __name__ == "__main__":
    from gui import App
    app = App()
    app.mainloop()
