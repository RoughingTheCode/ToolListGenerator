import os
import subprocess
from pathlib import Path

from heidenhain.settings import open_json
from heidenhain.app_paths import (
    APP_DIR, PROGRAMME_DIR, PLATZ_TABELLE_DIR, WERKZEUG_LISTE_DIR, AUSGABE_DIR, EINSTELLUNG_DIR,
    ensure_dirs
)


def machine_download(machine, pgm_list, subfolder):
    ensure_dirs()
    clear_program_folder()

    if machine == "_keine_":
        return None

    settings = open_json()
    tnc_path = settings["global_settings"]["path"]
    machine_info = settings["machine_settings"][machine]
    if pgm_list:
        error_test = get_programs_from_machine(pgm_list=pgm_list, machine_info=machine_info, tnc_path=tnc_path, subfolder=subfolder)
        if error_test:
            return error_test
        

    path = Path(machine_info["place_table_path"])
    error_test = get_data(path=path, destination_dir=PLATZ_TABELLE_DIR, machine_info=machine_info, tnc_path=tnc_path)
    if error_test:
        return error_test

    path = Path(machine_info["tool_table_path"])
    error_test = get_data(path=path, destination_dir=WERKZEUG_LISTE_DIR, machine_info=machine_info, tnc_path=tnc_path)
    if error_test:
        return error_test

    return None


def list_files():
    ensure_dirs()
    return os.listdir(PROGRAMME_DIR)



def clear_program_folder():
    ensure_dirs()

    def clear_dir(folder: Path):
        if not folder.exists():
            folder.mkdir(parents=True, exist_ok=True)
            return
        for p in folder.iterdir():
            if p.is_file():
                try:
                    p.unlink()
                except PermissionError:
                    continue

    clear_dir(PROGRAMME_DIR)
    clear_dir(PLATZ_TABELLE_DIR)
    clear_dir(WERKZEUG_LISTE_DIR)
    clear_dir(AUSGABE_DIR)


def get_programs_from_machine(pgm_list, machine_info, tnc_path, subfolder):
    if not machine_info:
        return {"error_title": "Fehler bei Maschinen Info",
                "error_message": "Programmierer melden"}
    
    for pgm in pgm_list:
        dest = PROGRAMME_DIR / pgm
        remote = get_remote(machine_info=machine_info, subfolder=subfolder, pgm=pgm)

        try:
            subprocess.run(
                [
                    tnc_path,
                    "/c",
                    "-i" + machine_info["ip_address"],
                    "GET",
                    remote,
                    str(dest),
                    "Exit",
                ],
                shell=False,
                stdout=True
            )
        except subprocess.SubprocessError:
            return {"error_title": "Fehler beim laden von Maschine",
                    "error_message": "Maschine nicht an?"}

    if not any(PROGRAMME_DIR.iterdir()):
        return {"error_title": "Fehler beim laden von Maschine",
                "error_message": "Programm existiert nicht"}

    return None


def get_data(path: Path, destination_dir: Path, machine_info, tnc_path):
    if path == Path(""):
        return None
    if not machine_info:
        return {"error_title": "Fehler bei Maschinen Info",
                "error_message": "Programmierer melden!"}
    destination_dir.mkdir(parents=True, exist_ok=True)

    try:
        subprocess.run(
            [
                tnc_path,
                "/c",
                "-i" + machine_info["ip_address"],
                "GET",
                path,
                str(destination_dir / "Liste"),
                "Exit",
            ],
            shell=False,
            stdout=True
        )
    except subprocess.SubprocessError:
        return {"error_title": "Fehler beim laden von Maschine",
                "error_message": "Maschine nicht an?"}

    # Zielordner leer?
    if not any(destination_dir.iterdir()):
        return {"error_title": "Fehler beim laden von Maschine",
                "error_message": "Programm existiert nicht"}

    return None

def get_remote(machine_info, subfolder, pgm):
    if subfolder != "_Bezugsordner_":
                remote = Path(machine_info["reference_folder"] + "\\" + subfolder + "\\" + pgm)
                return remote
    if machine_info["folder_structure"] == "Ja":
            start = int(machine_info["folder_name_start"]) - 1
            end = int(machine_info["folder_name_end"])
            folder = machine_info["reference_folder"]
            
            if folder.endswith("\\"): 
                remote = Path(folder + pgm[start:end] + "\\" + pgm)
                return remote
            else:
                remote = Path(folder + "\\" + pgm[start:end] + "\\" + pgm)
                return remote
    else:
        remote = Path(machine_info["reference_folder"] + pgm)
        return remote
