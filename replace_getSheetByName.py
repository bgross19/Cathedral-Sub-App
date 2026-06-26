import re

def main():
    with open('code.gs', 'r') as f:
        content = f.read()

    # The goal is to replace `ss.getSheetByName("Sheet Name")`
    # with `getSheetOrThrow(ss, "Sheet Name")`
    # or `sheetSS.getSheetByName("Sheet Name")`
    # with `getSheetOrThrow(sheetSS, "Sheet Name")`

    # Exclude definitions of getSheetOrThrow and getSheetCaseInsensitiveOrThrow
    # as they use getSheetByName internally

    # Let's do a simple regex replace and ignore lines that define the helper functions or use them inside the helper functions.

    new_content = ""
    lines = content.split('\n')
    inside_helper = False

    for line in lines:
        if line.startswith('function getSheetOrThrow') or line.startswith('function getSheetCaseInsensitiveOrThrow'):
            inside_helper = True
        elif inside_helper and line.startswith('}'):
            inside_helper = False
            new_content += line + "\n"
            continue

        if not inside_helper and 'getSheetByName(' in line:
            # We want to replace `<var_name>.getSheetByName(<string>)` with `getSheetOrThrow(<var_name>, <string>)`

            # Simple approach: find the variable before .getSheetByName
            # It could be `ss`, `sheetSS`, etc.
            match = re.search(r'([a-zA-Z0-9_]+)\.getSheetByName\((.*?)\)', line)
            if match:
                var_name = match.group(1)
                sheet_name = match.group(2)
                # print(f"Replacing {match.group(0)} with getSheetOrThrow({var_name}, {sheet_name})")
                new_line = line.replace(match.group(0), f"getSheetOrThrow({var_name}, {sheet_name})")
                new_content += new_line + "\n"
            else:
                new_content += line + "\n"
        else:
            new_content += line + "\n"

    with open('code.gs', 'w') as f:
        f.write(new_content)

if __name__ == "__main__":
    main()
