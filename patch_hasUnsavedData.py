import re

with open("Index.html", "r") as f:
    content = f.read()

# Make sure staffDutyInput is updated before hasUnsavedData checks it
# We can do this by updating staffDutyInput whenever a checkbox changes
search_pattern = r"""                <label class="flex items-center space-x-2"><input type="checkbox" value="1" class="form-checkbox h-4 w-4 text-\[#00843D\] staff-duty-cb"> <span class="text-sm">Period 1</span></label>"""

replacement = """                <label class="flex items-center space-x-2"><input type="checkbox" value="1" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb" onchange="updateHiddenDutyInput()"> <span class="text-sm">Period 1</span></label>"""

content = re.sub(search_pattern, replacement, content, count=1)

for i in range(2, 9):
    search_pattern = f"""                <label class="flex items-center space-x-2"><input type="checkbox" value="{i}" class="form-checkbox h-4 w-4 text-\\[#00843D\\] staff-duty-cb"> <span class="text-sm">Period {i}</span></label>"""
    replacement = f"""                <label class="flex items-center space-x-2"><input type="checkbox" value="{i}" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb" onchange="updateHiddenDutyInput()"> <span class="text-sm">Period {i}</span></label>"""
    content = re.sub(search_pattern, replacement, content, count=1)

# Add updateHiddenDutyInput function
search_pattern_func = r"""      function openAddStaffModal\(\) \{"""
replacement_func = """      function updateHiddenDutyInput() {
        let selectedDuties = [];
        document.querySelectorAll('.staff-duty-cb:checked').forEach(cb => selectedDuties.push(cb.value));
        document.getElementById('staffDutyInput').value = selectedDuties.join(', ');
      }

      function openAddStaffModal() {"""

content = re.sub(search_pattern_func, replacement_func, content, count=1)

with open("Index.html", "w") as f:
    f.write(content)

print("Success")
