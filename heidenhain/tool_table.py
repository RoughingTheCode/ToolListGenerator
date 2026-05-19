from heidenhain.app_paths import WERKZEUG_LISTE_DIR


def add_tool_values(tool_list, machine):
    data_files = sorted([p for p in WERKZEUG_LISTE_DIR.iterdir() if p.is_file()])
    if len(data_files) != 1:
        # Keine Werkzeugliste vorhanden
        return tool_list

    with open(data_files[0], "r", encoding="utf-8", errors="ignore") as f:
        file_lines = f.read().split("\n")
        file_head = file_lines[1:2]
        head_list = file_head[0].split()

        counter = 0
        for value in head_list:
            if value == "R":
                r = counter
            if value == "L":
                l = counter
            counter += 1

        file_lines = file_lines[2:-2]

    for primary_dict in tool_list:
        for tools in primary_dict.values():
            for tool_dict in tools:
                tool_key = list(tool_dict.keys())[0]
                if tool_dict[tool_key]["inside"]:
                    for line in file_lines:
                        line = line.split()
                        if not line:
                            continue
                        if str(tool_key) == line[0]:
                            tool_dict[tool_key]["length"] = line[l]
                            tool_dict[tool_key]["radius"] = line[r]

    return tool_list
