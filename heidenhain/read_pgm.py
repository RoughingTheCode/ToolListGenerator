import os
from heidenhain.app_paths import PROGRAMME_DIR
from heidenhain.settings import open_json


def pgm(tool_sort, clamp_info, machine):
    file_list = list_files()
    tool_list = []
    head_list = []

    for file in file_list:
        with open(PROGRAMME_DIR / file, "r", encoding="utf-8", errors="ignore") as file_text:
            pgm_text = file_text.read().split("\n")
            head = file_head(pgm_text, clamp_info=clamp_info)
            head_list.append({file[:-2]: head})

        settings = open_json()
        ikz_functions = settings["machine_settings"][machine]["ikz_functions"]

        tools = write_tool_list(pgm_text=pgm_text, tool_sort=tool_sort, ikz_functions=ikz_functions)
        tool_list.append({file[:-2]: tools})

    return tool_list, head_list

def pgm_with_path(tool_sort, clamp_info, pgm_path, machine):
    tool_list = []
    head_list = []
    pgm_name = pgm_path.split("/")
    pgm_name = pgm_name[-1][:-2]

    with open(pgm_path, "r", encoding="utf-8", errors="ignore") as file_text:
            pgm_text = file_text.read().split("\n")
            head = file_head(pgm_text, clamp_info=clamp_info)
            head_list.append({pgm_name: head})

    settings = open_json()
    ikz_functions = settings["machine_settings"][machine]["ikz_functions"]

    tools = write_tool_list(pgm_text=pgm_text, tool_sort=tool_sort, ikz_functions=ikz_functions)
    tool_list.append({pgm_name: tools})

    return tool_list, head_list


def list_files():
    if not PROGRAMME_DIR.exists():
        return []
    return os.listdir(PROGRAMME_DIR)


def file_head(pgm_text, clamp_info):
    head = []
    new_line = False
    cam = False
    counter = 0

    for line in pgm_text:
        counter += 1

        line = line_configure(line)

        if line.startswith("* - T"):
            return head

        elif new_line:
            new_line, line = test_new_line(line)
            head.append(line)

        elif line.startswith("*") and counter < 10:
            new_line, line = test_new_line(line)
            if line[4:].strip() != "":
                head.append(line[4:])
        elif clamp_info and line.startswith(";"):
            if not cam and len(line) > 3:
                new_line, line = test_new_line(line)
                line = line[1:].lstrip()
                head.append(line)
        if "-------------------------------------" in line:
                cam = True
        if cam and "Erzeugt am" in line or "Datei" in line:
                new_line, line = test_new_line(line)
                line = line[1:].lstrip()
                head.append(line)

    return head


def write_tool_list(pgm_text, tool_sort, ikz_functions):
    tools = []
    new_line = False
    cooling = None
    tool = None
    counter = 0

    for line in pgm_text:

        if "* - SICHERUNG" in line or "* - Sicherung" in line or "* - sicherung" in line:
            break

        if "* - T" in line or "*  - T" in line:

            if tool:
                if not cooling:
                    cooling = "A"
                tool = make_dict(tool=tool, cooling=cooling)
                if tool is not None:
                    tools.append(tool)

            counter = 0
            cooling = None

            if line.startswith("/"):
                line = line[1:]
            line = line_configure(line)
            line = line[5:]
            line = test_old_pgm(line)
            new_line, line = test_new_line(line)
            tool = line
            continue

        if new_line:
            new_line, line = test_new_line(line)
            line = line.lstrip().rstrip()
            if new_line:
                line = line[:-2]
                tool += line
            else:
                new_line = False
                tool += line

        if tool:
            counter += 1
            if counter <= 5:
                tool, new_line = cam_description(tool=tool, line=line)

            if counter <= 30:
                for function in ikz_functions:
                    if function in line:
                        cooling = "I"

    if tool:
        if not cooling:
            cooling = "A"
        tool = make_dict(tool=tool, cooling=cooling)
        tools.append(tool)

    if not tools:
        return None

    tools = sort_and_deduplicate_tools(tools, tool_sort)
    return tools


def test_new_line(line):
    if line.endswith("~"):
        line = line[:-1]
        return True, line
    return False, line


def line_configure(line):
    line = line.split()
    if line == []:
        return " "
    del line[0]
    line = " ".join(line)
    return line


def test_old_pgm(line):
    if line.startswith("-"):
        line = line[1:]
        line = line.split()
        line_tool = line[0].replace("-", " ")
        line = line[1:]
        line = " ".join(line)
        line = line_tool + " " + line

    elif "-" in line[:5]:
        line = line.split()
        line_tool = line[0].replace("-", " ")
        line = line[1:]
        line = " ".join(line)
        line = line_tool + " " + line

    return line


def sort_and_deduplicate_tools(tool_list, tool_sort):
    unique_tools_map = {}
    for item in tool_list:
        tool_number = list(item.keys())[0]
        tool_data = list(item.values())[0]

        if tool_number not in unique_tools_map:
            unique_tools_map[tool_number] = tool_data

    final_tool_list = []
    keys_to_iterate = sorted(unique_tools_map.keys()) if tool_sort else unique_tools_map.keys()

    for tool_number in keys_to_iterate:
        tool_data = unique_tools_map[tool_number]
        final_tool_list.append({tool_number: tool_data})

    return final_tool_list


def make_dict(tool, cooling):
    tool_split = tool.split(" ")
    tool_dict = None
    try:
        tool_number = int(tool_split[0])
    except ValueError:
        return None
    description = " ".join(tool_split[1:])
    if description.startswith("-"):
        description = description[2:]
    if tool_number:
        tool_dict = {tool_number: {"description": description,
                                   "length": None,
                                   "radius": None,
                                   "cooling": cooling,
                                   "inside": False}}
    return tool_dict


def cam_description(tool, line):
    line = line.split()
    new_line = False
    if line[0].isnumeric():
        line = line[1:]
        line = line = " ".join(line)
        if line.startswith(";"):
            line = line[1:]
            if line == "":
                return tool, new_line
            if line.endswith("~"):
                new_line = True
                line = line[:-1] + " "
            tool = tool + "\n" + line
            return tool, new_line
    else:
        return tool, new_line
    return tool, new_line
