import json
import os


SETTINGS_PATH = os.path.join("Einstellungen", "settings.json")


def save_json(dictionary: dict):
    os.makedirs("Einstellungen", exist_ok=True)

    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(dictionary, f, indent=4, ensure_ascii=False)


def open_json():
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            dictionary = json.load(f)
        return dictionary
    except FileNotFoundError:
        return None
    except json.JSONDecodeError:
        return None