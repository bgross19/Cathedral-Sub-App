import re

with open("Index.html", "r") as f:
    content = f.read()

# Fix HR logic
search_pattern1 = r"let teacherDutyArr = teacherDuty \? teacherDuty\.split\(\',\'\)\.map\(d => d\.trim\(\)\) : \[\];"
replacement1 = r"let teacherDutyArr = teacherDuty ? String(teacherDuty).split(',').map(d => d.trim()) : [];"

content = re.sub(search_pattern1, replacement1, content, count=1)

# Fix Stale State
search_pattern2 = r"""                      showToast\("Duty updated successfully.", "success"\);
                      // Optimistic UI, no need to reload entire list if we don't want to
                      // loadStaffRosterSettings\(\);"""

replacement2 = """                      showToast("Duty updated successfully.", "success");
                      loadStaffRosterSettings();"""

content = re.sub(search_pattern2, replacement2, content, count=1)

with open("Index.html", "w") as f:
    f.write(content)

print("Success")
