import json
from dataclasses import asdict, dataclass
from pathlib import Path


APP_HOME = Path.home() / ".reflow"
CONFIG_FILE = APP_HOME / "config.json"


@dataclass
class AppConfig:
    remove_markdown: bool = True
    normalize_punctuation: bool = True
    normalize_whitespace: bool = True
    merge_lines: bool = True
    remove_emoji: bool = False
    indent_paragraph: bool = True
    keep_tables: bool = True
    last_open_dir: str = ""
    last_save_dir: str = ""
    window_geometry: str = "1220x780"
    theme_mode: str = "light"


def load_config() -> AppConfig:
    if not CONFIG_FILE.exists():
        return AppConfig()
    try:
        raw = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return AppConfig()

    default = asdict(AppConfig())
    merged = {key: raw.get(key, value) for key, value in default.items()}
    return AppConfig(**merged)


def save_config(config: AppConfig) -> None:
    APP_HOME.mkdir(parents=True, exist_ok=True)
    CONFIG_FILE.write_text(json.dumps(asdict(config), ensure_ascii=False, indent=2), encoding="utf-8")
