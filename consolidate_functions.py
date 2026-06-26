import re

def main():
    with open('code.gs', 'r') as f:
        content = f.read()

    # The new refreshData function
    refresh_data_func = """
/**
 * Refreshes requested data components.
 * @param {Array<string>} components - The components to fetch (e.g. ['myAbsences', 'quickCover'])
 * @returns {Object} The requested data.
 */
function refreshData(components) {
  try {
    var payload = getInitialPayload();
    var response = {};
    for (var i = 0; i < components.length; i++) {
        var comp = components[i];
        if (payload[comp] !== undefined) {
            response[comp] = payload[comp];
        }
    }
    return response;
  } catch (e) {
    notifyAdminOfError("refreshData", e);
    throw new Error("Failed to refresh data: " + e.message);
  }
}
"""

    if "function refreshData" not in content:
        content += "\n" + refresh_data_func

    # Now we want to remove the obsolete backend functions
    funcs_to_remove = [
        "getStaffList",
        "getTodaysOpenJobsData",
        "getQuickCoverData",
        "getAdminDashboardData",
        "getMySubDuties",
        "getMyAbsences",
        "getHRDashboardData"
    ]

    for func in funcs_to_remove:
        # Regex to match function definition until the outermost closing brace.
        # This can be tricky with simple regex if there are nested braces.
        # Let's do a simple approach: find line with "function funcName()", count braces until we hit 0.
        lines = content.split('\n')
        new_lines = []
        inside_target = False
        brace_count = 0

        for line in lines:
            if re.match(rf'^function\s+{func}\s*\(', line):
                inside_target = True
                brace_count = line.count('{') - line.count('}')
                continue

            if inside_target:
                brace_count += line.count('{') - line.count('}')
                if brace_count <= 0:
                    inside_target = False
                continue

            new_lines.append(line)

        content = '\n'.join(new_lines)

    with open('code.gs', 'w') as f:
        f.write(content)

if __name__ == "__main__":
    main()
