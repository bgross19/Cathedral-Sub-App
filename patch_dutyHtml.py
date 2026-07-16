import re

with open("Index.html", "r") as f:
    content = f.read()

# Replace the dutyHtml generation
search_pattern = r"""          let dutyOptionsHtml = '<option value=""></option>';
          for \(let i = 1; i <= 8; i\+\+\) \{
              let isSelected = \(staff\.duty == i\) \? 'selected' : '';
              dutyOptionsHtml \+= `<option value="\$\{i\}" \$\{isSelected\}>\$\{i\}</option>`;
          \}

          let dutyHtml = `
              <select onchange="updateStaffDutyInline\('\$\{escapeJs\(staff\.email\)\}', this\.value\)" class="p-1 border rounded focus:ring-1 focus:ring-\[#00843D\] outline-none text-sm bg-white">
                  \$\{dutyOptionsHtml\}
              </select>
          `;"""

replacement = """          let dutyArr = staff.duty ? String(staff.duty).split(',').map(d => d.trim()) : [];
          let dutyDisplay = dutyArr.length > 0 ? dutyArr.join(', ') : 'None';

          let dutyOptionsHtml = '';
          for (let i = 1; i <= 8; i++) {
              let isChecked = dutyArr.includes(String(i)) ? 'checked' : '';
              dutyOptionsHtml += `
                  <label class="flex items-center space-x-2 p-1 hover:bg-gray-100 cursor-pointer text-sm">
                      <input type="checkbox" value="${i}" ${isChecked} onchange="handleDutyCheckboxChange('${escapeJs(staff.email)}', this)" class="form-checkbox h-4 w-4 text-[#00843D]">
                      <span>Period ${i}</span>
                  </label>
              `;
          }

          let dutyDropdownId = 'duty-dropdown-' + staff.email.replace(/[^a-zA-Z0-9]/g, '-');
          let dutyHtml = `
              <div class="relative inline-block text-left duty-dropdown-container">
                  <button type="button" onclick="toggleDutyDropdown('${dutyDropdownId}')" class="inline-flex justify-between w-full p-1 border rounded focus:ring-1 focus:ring-[#00843D] outline-none text-sm bg-white text-gray-700 hover:bg-gray-50">
                      <span id="duty-display-${dutyDropdownId}">${dutyDisplay}</span>
                      <svg class="-mr-1 ml-2 h-5 w-5" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" aria-hidden="true">
                          <path fill-rule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clip-rule="evenodd" />
                      </svg>
                  </button>
                  <div id="${dutyDropdownId}" class="hidden origin-top-right absolute right-0 mt-1 w-32 rounded-md shadow-lg bg-white ring-1 ring-black ring-opacity-5 z-10">
                      <div class="py-1" role="menu" aria-orientation="vertical">
                          ${dutyOptionsHtml}
                      </div>
                  </div>
              </div>
          `;"""

new_content = re.sub(search_pattern, replacement, content, count=1)
if new_content == content:
    print("No changes made. Pattern not found.")
else:
    with open("Index.html", "w") as f:
        f.write(new_content)
    print("Success")
