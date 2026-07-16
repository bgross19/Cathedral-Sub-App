import re

with open("Index.html", "r") as f:
    content = f.read()

# 1. Replace the input element with a grid of checkboxes
search_pattern1 = r"""            <div class="mb-4">
              <label class="block text-sm font-bold mb-1">Duty Number \(Optional\)</label>
              <input type="number" id="staffDutyInput" class="w-full p-2 border rounded-lg focus:ring-2 focus:ring-\[#00843D\] outline-none" placeholder="e\.g\. 1">
            </div>"""

replacement1 = """            <div class="mb-4">
              <label class="block text-sm font-bold mb-1">Duty Periods (Optional)</label>
              <input type="hidden" id="staffDutyInput" value="">
              <div class="grid grid-cols-4 gap-2 mt-2" id="staffDutyCheckboxes">
                <!-- Checkboxes will be dynamically generated or hardcoded here -->
                <label class="flex items-center space-x-2"><input type="checkbox" value="1" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 1</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="2" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 2</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="3" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 3</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="4" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 4</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="5" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 5</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="6" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 6</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="7" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 7</span></label>
                <label class="flex items-center space-x-2"><input type="checkbox" value="8" class="form-checkbox h-4 w-4 text-[#00843D] staff-duty-cb"> <span class="text-sm">Period 8</span></label>
              </div>
            </div>"""

content = re.sub(search_pattern1, replacement1, content, count=1)

# 2. Update saveStaffMember to gather checked values
search_pattern2 = r"""      function saveStaffMember\(\) \{
        const originalEmail = document\.getElementById\('staffOriginalEmail'\)\.value;
        const name = document\.getElementById\('staffNameInput'\)\.value;
        const email = document\.getElementById\('staffEmailInput'\)\.value;
        const role = document\.getElementById\('staffRoleInput'\)\.value;
        const duty = document\.getElementById\('staffDutyInput'\)\.value;"""

replacement2 = """      function saveStaffMember() {
        const originalEmail = document.getElementById('staffOriginalEmail').value;
        const name = document.getElementById('staffNameInput').value;
        const email = document.getElementById('staffEmailInput').value;
        const role = document.getElementById('staffRoleInput').value;

        let selectedDuties = [];
        document.querySelectorAll('.staff-duty-cb:checked').forEach(cb => selectedDuties.push(cb.value));
        const duty = selectedDuties.join(', ');

        // Keep hidden input in sync (optional, but good for captureModalState)
        document.getElementById('staffDutyInput').value = duty;"""

content = re.sub(search_pattern2, replacement2, content, count=1)

# 3. Update openAddStaffModal to uncheck all
search_pattern3 = r"""        document\.getElementById\('staffRoleInput'\)\.disabled = false;
        document\.getElementById\('staffDutyInput'\)\.value = "";"""

replacement3 = """        document.getElementById('staffRoleInput').disabled = false;
        document.getElementById('staffDutyInput').value = "";
        document.querySelectorAll('.staff-duty-cb').forEach(cb => cb.checked = false);"""

content = re.sub(search_pattern3, replacement3, content, count=1)

# 4. Update openEditStaffModal to check correct boxes
search_pattern4 = r"""        \} else \{
            roleInput\.value = \(normalizedRole\.toLowerCase\(\) === "substitute"\) \? "Substitute" : "Teacher";
            roleInput\.disabled = false;
        \}

        document\.getElementById\('staffDutyInput'\)\.value = duty \|\| "";"""

replacement4 = """        } else {
            roleInput.value = (normalizedRole.toLowerCase() === "substitute") ? "Substitute" : "Teacher";
            roleInput.disabled = false;
        }

        const dutyStr = duty || "";
        document.getElementById('staffDutyInput').value = dutyStr;
        const dutyArr = dutyStr.split(',').map(d => d.trim());
        document.querySelectorAll('.staff-duty-cb').forEach(cb => {
            cb.checked = dutyArr.includes(cb.value);
        });"""

content = re.sub(search_pattern4, replacement4, content, count=1)


with open("Index.html", "w") as f:
    f.write(content)

print("Success")
