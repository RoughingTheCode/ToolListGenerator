from heidenhain.settings import open_json


def remove_standard_tools(tool_list, machine):
    standard_tool_list = []
    settings = open_json()
    standard_tools = settings["machine_settings"][machine]["standard_tools"]

    for tool in standard_tools:
            if tool:
                standard_tool_list.append(int(tool))

    for primary_dict in tool_list:
        for key, tools in primary_dict.items():
            primary_dict[key] = [
                tool_dict
                for tool_dict in tools
                if next(iter(tool_dict)) not in standard_tool_list
            ]

    return tool_list
