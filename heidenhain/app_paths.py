# heidenhain/app_paths.py
import os
import sys
from pathlib import Path


def _find_root_from(start: Path) -> Path:
    """
    Suche nach einem Ordner, der typisch für dein Projekt ist (z.B. 'Bilder').
    Läuft ein paar Ebenen nach oben, bis gefunden.
    """
    start = start.resolve()
    for p in [start] + list(start.parents):
        if (p / "Bilder").is_dir():          # <- wichtigster Marker
            return p
    return start  # fallback


def app_dir() -> Path:
    # Wenn als EXE gestartet: Ordner der EXE
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    # Wenn als .py gestartet: heidenhain liegt unterhalb vom Projektroot
    here = Path(__file__).resolve().parent  # .../heidenhain
    return _find_root_from(here)


APP_DIR = app_dir()

BILDER_DIR = APP_DIR / "Bilder"
PROGRAMME_DIR = APP_DIR / "Programme"
PLATZ_TABELLE_DIR = APP_DIR / "Platz_Tabelle"
WERKZEUG_LISTE_DIR = APP_DIR / "Werkzeug_Liste"
AUSGABE_DIR = APP_DIR / "Ausgabeordner"
EINSTELLUNG_DIR = APP_DIR / "Einstellungen"


def ensure_dirs() -> None:
    PROGRAMME_DIR.mkdir(parents=True, exist_ok=True)
    PLATZ_TABELLE_DIR.mkdir(parents=True, exist_ok=True)
    WERKZEUG_LISTE_DIR.mkdir(parents=True, exist_ok=True)
    AUSGABE_DIR.mkdir(parents=True, exist_ok=True)
    EINSTELLUNG_DIR.mkdir(parents=True, exist_ok=True)


def chdir_to_app_dir() -> None:
    os.chdir(APP_DIR)
