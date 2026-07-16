import re

with open("Index.html", "r") as f:
    content = f.read()

# Add the JS functions toggleDutyDropdown and handleDutyCheckboxChange right after updateStaffDutyInline
search_pattern = r"""      function updateStaffDutyInline\(email, newDuty\) \{
          google\.script\.run
              \.withSuccessHandler\(result => \{
                  if \(result && result\.success\) \{
                      showToast\("Duty updated successfully.", "success"\);
                      loadStaffRosterSettings\(\);
                  \} else \{
                      showToast\("Failed to update duty: " \+ \(result \? result\.error : "Unknown error"\), "error"\);
                      loadStaffRosterSettings\(\); // revert
                  \}
              \}\)
              \.withFailureHandler\(err => \{
                  showToast\("Connection Error: " \+ err\.message, 'error'\);
                  loadStaffRosterSettings\(\); // revert
              \}\)
              \.updateStaffDutyInlineAdmin\(email, newDuty\);
      \}"""

replacement = """      function updateStaffDutyInline(email, newDuty) {
          google.script.run
              .withSuccessHandler(result => {
                  if (result && result.success) {
                      showToast("Duty updated successfully.", "success");
                      // Optimistic UI, no need to reload entire list if we don't want to
                      // loadStaffRosterSettings();
                  } else {
                      showToast("Failed to update duty: " + (result ? result.error : "Unknown error"), "error");
                      loadStaffRosterSettings(); // revert
                  }
              })
              .withFailureHandler(err => {
                  showToast("Connection Error: " + err.message, 'error');
                  loadStaffRosterSettings(); // revert
              })
              .updateStaffDutyInlineAdmin(email, newDuty);
      }

      function toggleDutyDropdown(dropdownId) {
          const dropdown = document.getElementById(dropdownId);
          if (dropdown.classList.contains('hidden')) {
              // Close all other dropdowns
              document.querySelectorAll('.duty-dropdown-container > div:not(.hidden)').forEach(el => {
                  el.classList.add('hidden');
              });
              dropdown.classList.remove('hidden');
          } else {
              dropdown.classList.add('hidden');
          }
      }

      function handleDutyCheckboxChange(email, checkbox) {
          // Find the container for this dropdown
          const dropdownContainer = checkbox.closest('.duty-dropdown-container');
          const checkboxes = dropdownContainer.querySelectorAll('input[type="checkbox"]');

          let selectedDuties = [];
          checkboxes.forEach(cb => {
              if (cb.checked) {
                  selectedDuties.push(cb.value);
              }
          });

          const newDutyStr = selectedDuties.join(', ');

          // Update display
          const dropdownId = dropdownContainer.querySelector('div[id^="duty-dropdown-"]').id;
          const displaySpan = document.getElementById('duty-display-' + dropdownId);
          displaySpan.innerText = newDutyStr || 'None';

          updateStaffDutyInline(email, newDutyStr);
      }

      // Close duty dropdowns when clicking outside
      document.addEventListener('click', function(event) {
          if (!event.target.closest('.duty-dropdown-container')) {
              document.querySelectorAll('.duty-dropdown-container > div:not(.hidden)').forEach(el => {
                  el.classList.add('hidden');
              });
          }
      });"""

new_content = re.sub(search_pattern, replacement, content, count=1)
if new_content == content:
    print("No changes made. Pattern not found.")
else:
    with open("Index.html", "w") as f:
        f.write(new_content)
    print("Success")
