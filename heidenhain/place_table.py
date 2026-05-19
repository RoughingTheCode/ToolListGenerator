from heidenhain.app_paths import PLATZ_TABELLE_DIR


def tool_in_machin(tools):
    tool_in_machine_list = []

    files = sorted([p for p in PLATZ_TABELLE_DIR.iterdir() if p.is_file()])
    if not files:
        # Keine Platz-Tabelle vorhanden
        return tools

    with open(files[0], "r", encoding="utf-8", errors="ignore") as file_text:
        place_table = file_text.read().split("\n")
        place_table = place_table[2:-2]

    for line in place_table:
        line = line.split()
        if not line:
            continue
        if line[1].startswith("%") or line[1] == "0" or line[1] == "F":
            continue
        tool_in_machine_list.append(line[1])

    for primary_dict in tools:
        for tool_list in primary_dict.values():
            for tool_dict in tool_list:
                tool_key = list(tool_dict.keys())[0]
                if str(tool_key) in tool_in_machine_list:
                    tool_dict[tool_key]["inside"] = True

    return tools
