import re

with open("Index.html", "r") as f:
    content = f.read()

search_pattern = r"""                    let teacherDuty = dutyLookup\[subName\];
                    if \(teacherDuty && teacherDuty === coveredPeriod\) \{"""

replacement = """                    let teacherDuty = dutyLookup[subName];
                    let teacherDutyArr = teacherDuty ? teacherDuty.split(',').map(d => d.trim()) : [];
                    if (teacherDutyArr.includes(String(coveredPeriod))) {"""

new_content = re.sub(search_pattern, replacement, content, count=1)

if new_content == content:
    print("No changes made. Pattern not found.")
else:
    with open("Index.html", "w") as f:
        f.write(new_content)
    print("Success")
