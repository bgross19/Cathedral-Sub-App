
      let currentAbsenceReasons = [];

      function loadSettingsData() {
        google.script.run
          .withSuccessHandler(settings => {
            const container = document.getElementById('generalSettingsContainer');

            try {
               currentAbsenceReasons = JSON.parse(settings["Absence Reasons"] || "[]");
            } catch(e) {
               currentAbsenceReasons = [];
            }

            let html = `
              <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label class="block text-sm font-bold mb-1">Email Mode</label>
                  <select id="settingsEmailMode" class="w-full p-2 border rounded-lg focus:ring-2 focus:ring-[#00843D] outline-none">
                    <option value="Live" ${settings["Email Mode"] === "Live" ? "selected" : ""}>Live (Normal Operation)</option>
                    <option value="Redirect" ${settings["Email Mode"] === "Redirect" ? "selected" : ""}>Redirect (Beta Testing)</option>
                    <option value="Off" ${settings["Email Mode"] === "Off" ? "selected" : ""}>Off (No Emails)</option>
                  </select>
                </div>
                <div>
                  <label class="block text-sm font-bold mb-1">Redirect Email Address</label>
                  <input type="email" id="settingsRedirectEmail" value="${escapeHtml(settings["Redirect Email"] || "")}" class="w-full p-2 border rounded-lg focus:ring-2 focus:ring-[#00843D] outline-none">
                </div>
              </div>
              <div class="mt-4 grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label class="block text-sm font-bold mb-1">App URL</label>
                  <input type="text" id="settingsAppUrl" value="${escapeHtml(settings["App URL"] || "")}" class="w-full p-2 border rounded-lg focus:ring-2 focus:ring-[#00843D] outline-none">
                </div>
                <div>
                  <label class="block text-sm font-bold mb-1">Urgency Cutoff Time (24h format)</label>
                  <input type="number" id="settingsUrgencyCutoff" min="0" max="23" value="${escapeHtml(settings["Urgency Cutoff Time"] || "15")}" class="w-full p-2 border rounded-lg focus:ring-2 focus:ring-[#00843D] outline-none" title="Hour of the day after which next-day absences become urgent. (e.g. 15 = 3 PM)">
                </div>
              </div>

              <div class="mt-6 border-t pt-4">
                 <div class="flex justify-between items-center mb-2">
                    <label class="block text-sm font-bold">Absence Reasons</label>
                    <button onclick="addAbsenceReasonRow()" class="bg-gray-200 hover:bg-gray-300 text-gray-800 text-xs font-bold py-1 px-3 rounded transition-colors">Add Reason</button>
                 </div>
                 <div id="absenceReasonsContainer" class="space-y-2 max-h-60 overflow-y-auto pr-2">
                 </div>
              </div>

              <div class="mt-6 flex justify-end">
                <button id="saveSettingsBtn" onclick="saveSettings()" class="bg-[#002147] hover:bg-blue-900 text-white font-bold py-2 px-6 rounded-lg transition-colors shadow-md">
                  Save General Settings
                </button>
              </div>
            `;
            container.innerHTML = html;
            renderAbsenceReasons();
          })
          .withFailureHandler(err => {
            document.getElementById('generalSettingsContainer').innerHTML = `<p class="text-red-600 font-bold">Failed to load settings: ${err.message}</p>`;
          })
          .getSettingsForFrontend();
      }

      function renderAbsenceReasons() {
          const container = document.getElementById('absenceReasonsContainer');
          if (!container) return;

          let html = '';
          if (currentAbsenceReasons.length === 0) {
              html = '<p class="text-gray-400 italic text-sm">No reasons configured.</p>';
          } else {
              currentAbsenceReasons.forEach((r, idx) => {
                  html += `
                    <div class="flex items-center gap-2 bg-gray-50 p-2 border rounded">
                        <input type="text" id="reason-text-${idx}" value="${escapeHtml(r.reason)}" placeholder="Reason name" class="flex-1 p-1 border rounded text-sm outline-none focus:ring-1 focus:ring-[#00843D]">
                        <label class="flex items-center gap-1 text-sm cursor-pointer ml-2">
                            <input type="checkbox" id="reason-hr-${idx}" ${r.hrRequired ? 'checked' : ''} class="h-4 w-4">
                            HR Req.
                        </label>
                        <button onclick="removeAbsenceReasonRow(${idx})" class="text-red-500 hover:text-red-700 font-bold px-2 ml-2" title="Remove">&times;</button>
                    </div>
                  `;
              });
          }
          container.innerHTML = html;
      }

      function addAbsenceReasonRow() {
          saveCurrentAbsenceReasonsState();
          currentAbsenceReasons.push({reason: "New Reason", hrRequired: false});
          renderAbsenceReasons();
      }

      function removeAbsenceReasonRow(idx) {
          saveCurrentAbsenceReasonsState();
          currentAbsenceReasons.splice(idx, 1);
          renderAbsenceReasons();
      }

      function saveCurrentAbsenceReasonsState() {
          const container = document.getElementById('absenceReasonsContainer');
          if (!container) return;

          for (let i = 0; i < currentAbsenceReasons.length; i++) {
              const textEl = document.getElementById(`reason-text-${i}`);
              const hrEl = document.getElementById(`reason-hr-${i}`);
              if (textEl && hrEl) {
                  currentAbsenceReasons[i].reason = textEl.value;
                  currentAbsenceReasons[i].hrRequired = hrEl.checked;
              }
          }
      }

      function saveSettings() {
        const btn = document.getElementById('saveSettingsBtn');
        const originalText = btn.innerText;
        btn.disabled = true;
        btn.innerText = "Saving...";

        saveCurrentAbsenceReasonsState();

        const newSettings = {
          "Email Mode": document.getElementById('settingsEmailMode').value,
          "Redirect Email": document.getElementById('settingsRedirectEmail').value,
          "App URL": document.getElementById('settingsAppUrl').value,
          "Urgency Cutoff Time": document.getElementById('settingsUrgencyCutoff').value,
          "Absence Reasons": JSON.stringify(currentAbsenceReasons)
        };

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              alert("Settings saved successfully!");
            } else {
              alert("Failed to save settings: " + (result ? result.error : "Unknown error"));
            }
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .withFailureHandler(err => {
            alert("Connection Error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .updateSettings(newSettings);
      }

      let currentRoles = [];

      function loadRolesData() {
        google.script.run
          .withSuccessHandler(roles => {
            currentRoles = roles;
            renderRolesTable();
          })
          .withFailureHandler(err => {
            document.getElementById('roleManagementTableBody').innerHTML = `<tr><td colspan="3" class="p-8 text-center text-red-600 font-bold">Failed to load roles: ${err.message}</td></tr>`;
          })
          .getUserRoles();
      }

      function renderRolesTable() {
        const tbody = document.getElementById('roleManagementTableBody');
        if (currentRoles.length === 0) {
          tbody.innerHTML = '<tr><td colspan="3" class="p-8 text-center text-gray-400 italic">No user roles found.</td></tr>';
          return;
        }

        let html = '';
        currentRoles.forEach((r, idx) => {
          html += `
            <tr class="hover:bg-gray-50 transition-colors border-b border-gray-100 last:border-b-0">
              <td class="p-3 text-sm text-gray-800 font-medium" id="role-email-text-${idx}">${escapeHtml(r.email)}</td>
              <td class="p-3 text-sm text-gray-600" id="role-role-text-${idx}">${escapeHtml(r.role)}</td>
              <td class="p-3 text-right space-x-2">
                <button onclick="editRole(${idx})" id="role-edit-btn-${idx}" class="text-blue-600 hover:text-blue-800 font-semibold text-sm">Edit</button>
                <button onclick="deleteRole('${escapeJs(r.email)}')" class="text-red-600 hover:text-red-800 font-semibold text-sm">Delete</button>
              </td>
            </tr>
          `;
        });
        tbody.innerHTML = html;
      }

      function editRole(idx) {
        const r = currentRoles[idx];
        const emailTd = document.getElementById(`role-email-text-${idx}`);
        const roleTd = document.getElementById(`role-role-text-${idx}`);
        const editBtn = document.getElementById(`role-edit-btn-${idx}`);

        const options = ["Admin", "Sub Coordinator", "HR", "Principal"];
        let selectHtml = `<select id="edit-role-select-${idx}" class="p-1 border rounded focus:ring-2 focus:ring-[#00843D] outline-none">`;
        options.forEach(opt => {
          selectHtml += `<option value="${opt}" ${r.role.toLowerCase() === opt.toLowerCase() ? 'selected' : ''}>${opt}</option>`;
        });
        selectHtml += `</select>`;

        emailTd.innerHTML = `<input type="email" id="edit-role-email-${idx}" value="${escapeHtml(r.email)}" class="p-1 border rounded w-full focus:ring-2 focus:ring-[#00843D] outline-none">`;
        roleTd.innerHTML = selectHtml;

        editBtn.innerText = "Save";
        editBtn.onclick = function() { saveRoleEdit(idx, r.email); };
      }

      function saveRoleEdit(idx, oldEmail) {
        const newEmail = document.getElementById(`edit-role-email-${idx}`).value;
        const newRole = document.getElementById(`edit-role-select-${idx}`).value;

        if (!newEmail || !newRole) {
          alert("Email and Role cannot be empty.");
          return;
        }

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              loadRolesData(); // Reload table
            } else {
              alert("Failed to update role: " + (result ? result.error : "Unknown error"));
              loadRolesData(); // Revert on failure
            }
          })
          .withFailureHandler(err => {
            alert("Connection Error: " + err.message);
            loadRolesData();
          })
          .editUserRole(oldEmail, newEmail, newRole);
      }

      function deleteRole(email) {
        if (!confirm(`Are you sure you want to remove the role for ${email}?`)) return;

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              loadRolesData();
            } else {
              alert("Failed to delete role: " + (result ? result.error : "Unknown error"));
            }
          })
          .withFailureHandler(err => {
            alert("Connection Error: " + err.message);
          })
          .deleteUserRole(email);
      }

      function openAddRoleModal() {
        document.getElementById('addRoleEmail').value = '';
        document.getElementById('addRoleSelect').value = 'Admin';

        const modal = document.getElementById('addRoleModal');
        const overlay = document.getElementById('genericModalOverlay');
        const body = document.querySelector('body');

        modal.classList.remove('opacity-0');
        modal.classList.remove('pointer-events-none');
        overlay.classList.remove('opacity-0');
        overlay.classList.remove('pointer-events-none');
        body.classList.add('modal-active');
        // Delay slightly to let inputs initialize if needed
        setTimeout(captureModalState, 50);
      }

      function closeAddRoleModal() {
        const modal = document.getElementById('addRoleModal');
        const overlay = document.getElementById('genericModalOverlay');
        const body = document.querySelector('body');

        modal.classList.add('opacity-0');
        modal.classList.add('pointer-events-none');
        overlay.classList.add('opacity-0');
        overlay.classList.add('pointer-events-none');
        body.classList.remove('modal-active');
      }

      function submitAddRole() {
        const email = document.getElementById('addRoleEmail').value;
        const role = document.getElementById('addRoleSelect').value;

        if (!email || !role) {
          alert("Please fill in both email and role.");
          return;
        }

        const btn = document.getElementById('addRoleSubmitBtn');
        const originalText = btn.innerText;
        btn.disabled = true;
        btn.innerText = "Adding...";

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              closeAddRoleModal();
              loadRolesData();
            } else {
              alert("Failed to add role: " + (result ? result.error : "Unknown error"));
            }
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .withFailureHandler(err => {
            alert("Connection Error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .addUserRole(email, role);
      }


      function toggleSection(wrapperId, iconId) {
        const wrapper = document.getElementById(wrapperId);
        const icon = document.getElementById(iconId);

        // Check current state
        const isOpen = wrapper.classList.contains('grid-rows-1');

        if (isOpen) {
          // Close it
          wrapper.classList.remove('grid-rows-1');
          wrapper.classList.add('grid-rows-0');
          icon.classList.add('-rotate-90');
          localStorage.setItem(`toggle_${wrapperId}`, 'closed');
        } else {
          // Open it
          wrapper.classList.remove('grid-rows-0');
          wrapper.classList.add('grid-rows-1');
          icon.classList.remove('-rotate-90');
          localStorage.setItem(`toggle_${wrapperId}`, 'open');
        }
      }

      function initToggles() {
        const sections = [
          { wrapperId: 'myAbsencesWrapper', iconId: 'iconMyAbsences' },
          { wrapperId: 'mySubDutiesWrapper', iconId: 'iconMySubDuties' },
          { wrapperId: 'todaysOpenJobsWrapper', iconId: 'iconTodaysOpenJobs' }
        ];

        sections.forEach(sec => {
          const state = localStorage.getItem(`toggle_${sec.wrapperId}`);
          if (state === 'closed') {
            const wrapper = document.getElementById(sec.wrapperId);
            const icon = document.getElementById(sec.iconId);
            if (wrapper && icon) {
              wrapper.classList.remove('grid-rows-1');
              wrapper.classList.add('grid-rows-0');
              icon.classList.add('-rotate-90');
            }
          }
        });
      }

      let staffList = [];
      let currentUnfilledData = [];
      let myAbsencesData = [];
      let mySubDutiesData = [];
      function escapeHtml(str) {
        if (str === null || str === undefined) return '';
        if (str === 0) return '0';
        if (str === false) return 'false';
        if (str === '') return '';
        return String(str)
          .replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;')
          .replace(/'/g, '&#39;');
      }

      function escapeJs(str) {
        if (str === null || str === undefined) return '';
        const jsEscaped = String(str)
          .replace(/\\/g, '\\\\')
          .replace(/'/g, "\\'")
          .replace(/"/g, '\\"')
          .replace(/\n/g, '\\n')
          .replace(/\r/g, '\\r');
        return escapeHtml(jsEscaped);
      }

      let adminDataRaw = [];
      let adminDataFiltered = [];
      let adminSortCol = 'date';
      let adminSortAsc = true;
      let adminCurrentView = 'period'; // 'period' or 'request'
      let hrDataRaw = [];
      let hrPayPeriods = [];
      let hrCurrentView = 'absences'; // 'absences' or 'subs'

      function initApp() {
        initToggles();
        google.script.run
          .withSuccessHandler(payload => {
            const data = payload.userData;
            document.getElementById('userName').innerText = data.name;
            document.getElementById('welcomeName').innerText = (data.name.split(',')[1] || data.name).trim();

            if (data.appUrl) {
                const footerLink = document.getElementById('atAGlanceFooterUrl');
                if (footerLink) footerLink.href = data.appUrl;
            }

            try {
               let reasons = JSON.parse(data.absenceReasons || "[]");
               let html = "";
               reasons.forEach(r => {
                 // Store hrRequired property dynamically to be accessed later
                 html += `<option value="${escapeHtml(r.reason)}" data-hr-required="${r.hrRequired}">${escapeHtml(r.reason)}</option>`;
               });
               if (html === "") {
                 html = `<option value="Personal" data-hr-required="false">Personal</option>`; // Fallback
               }
               document.getElementById('absReason').innerHTML = html;
               checkSpecialReason(); // Call initially to set HR checkbox state
            } catch(e) {
               console.error("Failed to parse absence reasons", e);
               document.getElementById('absReason').innerHTML = `<option value="Personal" data-hr-required="false">Personal</option>`;
            }

            // Render common parts
            renderMyAbsences(payload.myAbsences);
            renderMySubDuties(payload.mySubDuties);
            renderTodaysOpenJobs(payload.todaysOpenJobs);

            let lowerRole = data.role.toLowerCase();
            if (lowerRole === 'admin' || lowerRole === 'sub coordinator' || lowerRole === 'hr' || lowerRole === 'principal') {
               if (payload.staffList) renderStaffList(payload.staffList);
            }

            if (lowerRole === 'admin' || lowerRole === 'sub coordinator') {
              document.getElementById('adminSection').classList.remove('hidden');
              document.getElementById('navAdminBtn').classList.remove('hidden');

              if (payload.adminData) {
                  renderAdminDataRawSilent(payload.adminData);
              }
              if (payload.quickCover) {
                  renderQuickCover(payload.quickCover);
              }
            }
            if (lowerRole === 'admin') {
              document.getElementById('navSettingsBtn').classList.remove('hidden');
            }
            if (lowerRole === 'sub coordinator') {
              document.getElementById('navAtAGlanceBtn').classList.remove('hidden');
            }
            if (lowerRole === 'hr' || lowerRole === 'principal') {
              document.getElementById('navHRBtn').classList.remove('hidden');
              if (payload.hrData) {
                  hrDataRaw = payload.hrData.requests || [];
                  hrPayPeriods = payload.hrData.payPeriods || [];
              }
            }
            if (lowerRole === 'hr') {
              toggleHRDashboard();
              // Hide the return to homepage button inside the HR dashboard
              const returnBtn = document.querySelector('#hrDashboardView button[onclick="toggleHRDashboard()"]');
              if (returnBtn) returnBtn.classList.add('hidden');
            }
          })
          .withFailureHandler(err => {
            console.error(err);
            document.getElementById('userName').innerText = "Authentication Error";

            // Also notify parts that failed to load
            document.getElementById('myAbsences').innerHTML = `<div class="text-red-600 font-bold">Failed to load absences: ${err.message}</div>`;
            document.getElementById('mySubDuties').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load sub-duties: ${err.message}</div>`;
            document.getElementById('todaysOpenJobs').innerHTML = `<div class="text-red-600 font-bold">Failed to load open jobs: ${err.message}</div>`;
          })
          .getInitialPayload();
      }

      function renderStaffList(list) {
            staffList = list;
            // Rerender Quick Cover in case it loaded before the staff list
            if (typeof currentUnfilledData !== 'undefined' && currentUnfilledData && currentUnfilledData.length > 0) {
               renderQuickCover(currentUnfilledData);
            }
      }

      function loadStaffList() {
        google.script.run
          .withSuccessHandler(renderStaffList)
          .getStaffList();
      }


      // --- MODAL CLOSE AND UNSAVED DATA LOGIC ---
      let originalModalData = {};

      function captureModalState() {
        const isAddRoleModalOpen = !document.getElementById('addRoleModal').classList.contains('opacity-0');
        const isRequestModalOpen = !document.querySelector('#modal').classList.contains('opacity-0');

        if (isAddRoleModalOpen) {
          originalModalData = {
            email: document.getElementById('addRoleEmail').value,
            role: document.getElementById('addRoleSelect').value
          };
        } else if (isRequestModalOpen) {
          const periodCbs = document.querySelectorAll('.period-cb');
          const checkedPeriods = [];
          periodCbs.forEach(cb => { if(cb.checked) checkedPeriods.push(cb.value) });

          originalModalData = {
            date: document.getElementById('absDate').value,
            duration: document.getElementById('absDuration').value,
            reason: document.getElementById('absReason').value,
            urgency: document.getElementById('urgencyValue').value,
            specialInst: document.getElementById('specialInst').value,
            periods: checkedPeriods.join(',')
          };
        } else {
          originalModalData = {};
        }
      }

      function hasUnsavedData() {
        const isAddRoleModalOpen = !document.getElementById('addRoleModal').classList.contains('opacity-0');
        const isRequestModalOpen = !document.querySelector('#modal').classList.contains('opacity-0');

        if (isAddRoleModalOpen) {
          const currentEmail = document.getElementById('addRoleEmail').value;
          const currentRole = document.getElementById('addRoleSelect').value;

          if (originalModalData.email === undefined && currentEmail.trim() !== "") return true;
          if (originalModalData.email !== undefined && (currentEmail !== originalModalData.email || currentRole !== originalModalData.role)) return true;
        }

        if (isRequestModalOpen) {
          const periodCbs = document.querySelectorAll('.period-cb');
          const checkedPeriods = [];
          periodCbs.forEach(cb => { if(cb.checked) checkedPeriods.push(cb.value) });

          const currentDate = document.getElementById('absDate').value;
          const currentDuration = document.getElementById('absDuration').value;
          const currentReason = document.getElementById('absReason').value;
          const currentUrgency = document.getElementById('urgencyValue').value;
          const currentSpecialInst = document.getElementById('specialInst').value;
          const currentPeriods = checkedPeriods.join(',');

          const isEdit = document.getElementById('editAbsenceId').value !== "";

          if (!isEdit) {
            // Adding new: if they typed *anything* that isn't default
            if (currentDate || currentSpecialInst.trim() !== "" || checkedPeriods.length > 0) return true;
          } else {
            // Editing existing: compare to original state
            if (originalModalData.date !== undefined && (
                currentDate !== originalModalData.date ||
                currentDuration !== originalModalData.duration ||
                currentReason !== originalModalData.reason ||
                currentUrgency !== originalModalData.urgency ||
                currentSpecialInst !== originalModalData.specialInst ||
                currentPeriods !== originalModalData.periods)) {
                return true;
            }
          }
        }

        return false;
      }

      function attemptCloseAllModals() {
        if (hasUnsavedData()) {
          if (!confirm("You have unsaved changes. Are you sure you want to close this modal and discard them?")) {
            return;
          }
        }

        // Reset the captured state
        originalModalData = {};

        closeAllModals();
      }

      document.addEventListener('keydown', function(event) {
        if (event.key === "Escape") {
          const modals = document.querySelectorAll('.modal');
          let anyOpen = false;
          modals.forEach(m => {
            if (!m.classList.contains('opacity-0') && m.id !== 'genericModalOverlay') {
              anyOpen = true;
            }
          });

          if (anyOpen) {
            attemptCloseAllModals();
          }
        }
      });
      // --- END MODAL CLOSE AND UNSAVED DATA LOGIC ---

      function closeAllModals() {
        const modals = document.querySelectorAll('.modal');
        modals.forEach(m => {
          m.classList.add('opacity-0');
          m.classList.add('pointer-events-none');
        });
        document.querySelector('body').classList.remove('modal-active');
      }

      function toggleModal(isEdit = false, absId = null) {
        const body = document.querySelector('body');
        const modal = document.querySelector('#modal');
        const overlay = document.querySelector('#genericModalOverlay');

        if (modal.classList.contains('opacity-0')) {
          // Opening
          modal.classList.remove('opacity-0');
          modal.classList.remove('pointer-events-none');
          overlay.classList.remove('opacity-0');
          overlay.classList.remove('pointer-events-none');
          body.classList.add('modal-active');

          if (!isEdit) {
             document.getElementById('requestModalTitle').innerText = "Request Sub";
             document.getElementById('absenceForm').reset();
             document.getElementById('editAbsenceId').value = "";
             document.getElementById('submitBtn').innerText = "Submit Request";
             setUrgency('Standard');
             checkSpecialReason();
             document.getElementById('adminDeleteBtn').classList.add('hidden');

             // Uncheck all periods
             document.querySelectorAll('.period-cb').forEach(cb => cb.checked = false);
          }
          // Delay slightly to let inputs initialize if needed
          setTimeout(captureModalState, 50);
        } else {
          // Closing
          modal.classList.add('opacity-0');
          modal.classList.add('pointer-events-none');
          overlay.classList.add('opacity-0');
          overlay.classList.add('pointer-events-none');
          body.classList.remove('modal-active');
        }
      }

      function checkSpecialReason() {
        const select = document.getElementById('absReason');
        const hrSection = document.getElementById('hrSection');

        let hrRequired = false;
        if (select.options.length > 0 && select.selectedIndex > -1) {
            hrRequired = select.options[select.selectedIndex].getAttribute('data-hr-required') === 'true';
        }

        if (hrRequired) {
          hrSection.classList.remove('hidden');
        } else {
          hrSection.classList.add('hidden');
        }
      }

      function setUrgency(val) {
        document.getElementById('urgencyValue').value = val;
        const s = document.getElementById('btnStd');
        const u = document.getElementById('btnUrg');
        if (val === 'Standard') {
          s.className = "flex-1 py-2 rounded-lg border-2 border-[#00843D] bg-[#00843D] text-white font-bold transition-colors";
          u.className = "flex-1 py-2 rounded-lg border-2 border-gray-200 text-gray-500 font-bold transition-colors";
        } else {
          u.className = "flex-1 py-2 rounded-lg border-2 border-red-600 bg-red-600 text-white font-bold transition-colors";
          s.className = "flex-1 py-2 rounded-lg border-2 border-gray-200 text-gray-500 font-bold transition-colors";
        }
      }

      let todaysOpenJobsData = [];

      function renderTodaysOpenJobs(unfilled) {
            const container = document.getElementById('todaysOpenJobs');

            if (unfilled.length === 0) {
              container.innerHTML = 'There are no open jobs available for today.';
              container.classList.add('text-gray-400');
              return;
            }

            let html = '<div class="divide-y divide-gray-200 text-left">';
            todaysOpenJobsData = unfilled;

            unfilled.forEach((req, index) => {
              let assignId = `todays-assign-${req.id}-${req.period}`;
              let roomPillClass = req.room === "No Class Assigned" ? "bg-gray-100 text-gray-600 border-gray-200" : "bg-blue-100 text-blue-800 border-blue-200";
              let coursePillClass = req.course === "No Class Assigned" ? "bg-gray-100 text-gray-600 border-gray-200" : "bg-purple-100 text-purple-800 border-purple-200";

              html += `
                <div class="p-4 hover:bg-orange-50 transition-colors flex flex-col sm:flex-row justify-between items-start sm:items-center bg-white gap-4">
                  <div>
                    <p class="font-bold text-[#002147]">${escapeHtml(req.date)}</p>
                    <p class="text-sm text-gray-600 font-semibold">${escapeHtml(req.teacherName)}</p>

                    <div class="mt-2 flex flex-col gap-1 items-start">
                       <span class="inline-block bg-orange-100 text-orange-800 text-xs font-bold px-3 py-1 rounded-full border border-orange-200">Period ${escapeHtml(req.period)}</span>
                       <span class="inline-block ${roomPillClass} text-xs font-bold px-3 py-1 rounded-full border">Room: ${escapeHtml(req.room)}</span>
                       <span class="inline-block ${coursePillClass} text-xs font-bold px-3 py-1 rounded-full border truncate max-w-xs">Course: ${escapeHtml(req.course)}</span>
                    </div>
                  </div>
                  <div class="flex items-center gap-2 w-full sm:w-auto" onclick="event.stopPropagation()">
                    <button id="${escapeHtml(assignId)}" onclick="handleSignMeUp('${escapeJs(req.id)}', '${escapeJs(req.period)}', '${escapeJs(assignId)}')" class="bg-[#00843D] hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg text-sm transition-colors shadow-sm">
                      Sign Me Up
                    </button>
                  </div>
                </div>
              `;
            });
            html += '</div>';

            container.innerHTML = html;
            container.classList.remove('text-gray-400');
          }

      function loadTodaysOpenJobs() {
        google.script.run
          .withSuccessHandler(renderTodaysOpenJobs)
          .withFailureHandler(err => {
            document.getElementById('todaysOpenJobs').innerHTML = `<div class="text-red-600 font-bold">Failed to load open jobs: ${err.message}</div>`;
          })
          .getTodaysOpenJobsData();
      }

      function handleSignMeUp(absenceId, period, btnId) {
        if (!confirm("Are you sure you want to cover this period?")) return;

        const btn = document.getElementById(btnId);
        const originalText = btn.innerText;
        btn.disabled = true;
        btn.innerText = "Assigning...";

        // Get user name loaded in initApp
        const subName = document.getElementById('userName').innerText;

        if (!subName || subName === "Loading...") {
          alert("Error: User name not fully loaded yet. Please try again in a moment.");
          btn.disabled = false;
          btn.innerText = originalText;
          return;
        }

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              alert("Successfully signed up for period " + period + "!");
              // Refresh all relevant sections
              loadTodaysOpenJobs();
              loadMySubDuties();

              if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 loadQuickCover();
                 loadAdminDataRawSilent();
              }
            } else {
              alert("Failed to assign sub: " + (result ? result.error : "Unknown error"));
              btn.disabled = false;
              btn.innerText = originalText;
            }
          })
          .withFailureHandler(err => {
            alert("Connection error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .assignSubToPeriod(absenceId, period, subName);
      }

      function updateDynamicStaffList(dateStr, period) {
          const datalist = document.getElementById('dynamicStaffDataList');
          if (!datalist) return;

          // Find all subs currently busy at this exact date and period
          const busySubs = new Set();
          adminDataRaw.forEach(req => {
              if (req.formDateString === dateStr && String(req.period) === String(period)) {
                  if (req.assignedSub) {
                      busySubs.add(req.assignedSub.trim().toLowerCase());
                  }
              }
          });

          // Build HTML for subs who are NOT busy
          let optionsHtml = '';
          staffList.forEach(staff => {
              if (!busySubs.has(staff.name.trim().toLowerCase())) {
                  optionsHtml += `<option value="${escapeHtml(staff.display)}"></option>`;
              }
          });

          datalist.innerHTML = optionsHtml;
      }

      function renderQuickCover(unfilled) {
            const container = document.getElementById('quickCoverContainer');

            if (unfilled.length === 0) {
              container.innerHTML = '<div class="p-8 text-center text-gray-500 font-medium flex flex-col items-center gap-2"><span class="text-3xl">🎉</span> All classes are covered for the next 2 days!</div>';
              return;
            }

            let html = '<div class="divide-y divide-gray-200">';

            // Add a container for dynamic datalists at the top level of this section
            if (!document.getElementById('dynamicStaffDataList')) {
                const dynamicListEl = document.createElement('datalist');
                dynamicListEl.id = 'dynamicStaffDataList';
                document.body.appendChild(dynamicListEl);
            }

            currentUnfilledData = unfilled;
            unfilled.forEach((req, index) => {
              let assignId = `assign-${req.id}-${req.period}`;

              let roomPillClass = req.room === "No Class Assigned" ? "bg-gray-100 text-gray-600 border-gray-200" : "bg-blue-100 text-blue-800 border-blue-200";
              let coursePillClass = req.course === "No Class Assigned" ? "bg-gray-100 text-gray-600 border-gray-200" : "bg-purple-100 text-purple-800 border-purple-200";

              html += `
                <div onclick="openDetailsModal(${index})" class="p-4 hover:bg-orange-50 transition-colors flex flex-col sm:flex-row justify-between items-start sm:items-center bg-white gap-4 cursor-pointer">
                  <div>
                    <p class="font-bold text-[#002147]">${escapeHtml(req.date)}</p>
                    <p class="text-sm text-gray-600 font-semibold">${escapeHtml(req.teacherName)}</p>

                    <div class="mt-2 flex flex-col gap-1 items-start">
                       <span class="inline-block bg-orange-100 text-orange-800 text-xs font-bold px-3 py-1 rounded-full border border-orange-200">Period ${escapeHtml(req.period)}</span>
                       <span class="inline-block ${roomPillClass} text-xs font-bold px-3 py-1 rounded-full border">Room: ${escapeHtml(req.room)}</span>
                       <span class="inline-block ${coursePillClass} text-xs font-bold px-3 py-1 rounded-full border truncate max-w-xs">Course: ${escapeHtml(req.course)}</span>
                    </div>
                  </div>
                  <div class="flex items-center gap-2 w-full sm:w-auto" onclick="event.stopPropagation()">
                    <input type="text" id="${escapeHtml(assignId)}" list="dynamicStaffDataList" onfocus="updateDynamicStaffList('${escapeJs(req.formDateString)}', '${escapeJs(req.period)}')" placeholder="Enter sub name..." class="flex-1 sm:w-48 p-2 border rounded-lg text-sm outline-none focus:ring-2 focus:ring-[#00843D] bg-white">
                    <button onclick="handleAssign('${escapeJs(req.id)}', '${escapeJs(req.period)}', '${escapeJs(assignId)}')" class="bg-[#00843D] hover:bg-green-700 text-white font-bold py-2 px-4 rounded-lg text-sm transition-colors shadow-sm">
                      Assign
                    </button>
                  </div>
                </div>
              `;
            });
            html += '</div>';

            container.innerHTML = html;
          }

      function loadQuickCover() {
        google.script.run
          .withSuccessHandler(renderQuickCover)
          .withFailureHandler(err => {
            document.getElementById('quickCoverContainer').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load data: ${err.message}</div>`;
          })
          .getQuickCoverData();
      }

      function toggleSettingsDashboard() {
        const mainContent = document.getElementById('mainContent');
        const adminView = document.getElementById('adminDashboardView');
        const hrView = document.getElementById('hrDashboardView');
        const atAGlanceView = document.getElementById('atAGlanceDashboardView');
        const settingsView = document.getElementById('settingsDashboardView');

        if (settingsView.classList.contains('hidden')) {
          mainContent.classList.add('hidden');
          if (adminView) adminView.classList.add('hidden');
          if (hrView) hrView.classList.add('hidden');
          if (atAGlanceView) atAGlanceView.classList.add('hidden');
          settingsView.classList.remove('hidden');

          // load Settings Data
          loadSettingsData();
          loadRolesData();
        } else {
          settingsView.classList.add('hidden');
          mainContent.classList.remove('hidden');
        }
      }

      function toggleAdminDashboard() {
        const mainContent = document.getElementById('mainContent');
        const adminView = document.getElementById('adminDashboardView');
        const hrView = document.getElementById('hrDashboardView');
        const atAGlanceView = document.getElementById('atAGlanceDashboardView');
        const settingsView = document.getElementById('settingsDashboardView');

        if (adminView.classList.contains('hidden')) {
          mainContent.classList.add('hidden');
          if (hrView) hrView.classList.add('hidden');
          if (atAGlanceView) atAGlanceView.classList.add('hidden');
          if (settingsView) settingsView.classList.add('hidden');
          adminView.classList.remove('hidden');

          const today = new Date();
          const offset = today.getTimezoneOffset() * 60000;
          const localISOTime = (new Date(today - offset)).toISOString().split('T')[0];

          if (!document.getElementById('filterStartDate').value && !document.getElementById('filterEndDate').value) {
              document.getElementById('filterStartDate').value = localISOTime;
          }

          if (adminDataRaw.length === 0) {
            loadAdminDashboardData();
          } else {
            applyAdminFilters();
          }
        } else {
          adminView.classList.add('hidden');
          mainContent.classList.remove('hidden');
        }
      }

      function toggleAtAGlanceDashboard() {
        const mainContent = document.getElementById('mainContent');
        const adminView = document.getElementById('adminDashboardView');
        const hrView = document.getElementById('hrDashboardView');
        const atAGlanceView = document.getElementById('atAGlanceDashboardView');
        const settingsView = document.getElementById('settingsDashboardView');

        if (atAGlanceView.classList.contains('hidden')) {
          mainContent.classList.add('hidden');
          if (adminView) adminView.classList.add('hidden');
          if (hrView) hrView.classList.add('hidden');
          if (settingsView) settingsView.classList.add('hidden');
          atAGlanceView.classList.remove('hidden');

          if (adminDataRaw.length === 0) {
            loadAdminDashboardData(true); // pass flag to indicate it should render atAGlance when done
          } else {
            // Default to today if empty
            if (!document.getElementById('atAGlanceDate').value) {
                const today = new Date();
                const offset = today.getTimezoneOffset() * 60000;
                const localISOTime = (new Date(today - offset)).toISOString().split('T')[0];
                document.getElementById('atAGlanceDate').value = localISOTime;
            }
            applyAtAGlanceFilter();
          }
        } else {
          atAGlanceView.classList.add('hidden');
          mainContent.classList.remove('hidden');
        }
      }

      function applyAtAGlanceFilter() {
        const selectedDate = document.getElementById('atAGlanceDate').value;
        if (!selectedDate) return;

        // Filter for records on the exact selected date
        const todaysRecords = adminDataRaw.filter(req => req.date === selectedDate);
        renderAtAGlanceTable(todaysRecords);
      }

      function renderAtAGlanceTable(records) {
        const thead = document.getElementById('atAGlanceTableHeader');
        const tbody = document.getElementById('atAGlanceTableBody');

        if (records.length === 0) {
            thead.innerHTML = '<tr class="bg-gray-50 border-b border-gray-300 text-sm text-gray-700 uppercase"><th class="p-3 border border-gray-300">Teacher Name</th></tr>';
            tbody.innerHTML = '<tr><td class="p-8 text-center text-gray-400 italic border border-gray-300">No absence requests found for this date.</td></tr>';
            return;
        }

        // 1. Group records by teacher
        const teacherMap = {};
        const activePeriods = new Set();

        records.forEach(req => {
            if (!teacherMap[req.teacherName]) {
                teacherMap[req.teacherName] = {
                    name: req.teacherName,
                    periods: {}
                };
            }

            const rawIdx = adminDataRaw.findIndex(r => r.id === req.id && r.period === req.period);

            teacherMap[req.teacherName].periods[req.period] = {
                assignedSub: req.assignedSub || null,
                room: req.room || "N/A",
                rawIdx: rawIdx
            };
            activePeriods.add(parseInt(req.period));
        });

        const sortedActivePeriods = Array.from(activePeriods).sort((a, b) => a - b);

        // 2. Build Header
        let headerHtml = `<th class="p-3 border border-gray-300 min-w-[150px]">Teacher Name</th>`;
        sortedActivePeriods.forEach(p => {
            headerHtml += `<th class="p-3 border border-gray-300 text-center w-32">Period ${p}</th>`;
        });
        thead.innerHTML = headerHtml;

        // 3. Build Body
        let html = '';
        const teachers = Object.values(teacherMap).sort((a, b) => a.name.localeCompare(b.name));

        teachers.forEach(t => {
            html += `<tr class="hover:bg-gray-50 transition-colors">`;
            html += `<td class="p-3 text-sm text-gray-800 font-bold border border-gray-300">${t.name}</td>`;

            sortedActivePeriods.forEach(p => {
                const pData = t.periods[p];
                if (pData) {
                    // They requested a sub for this period
                    if (pData.assignedSub) {
                        html += `<td class="p-3 text-xs text-center border border-gray-300 bg-green-50 align-top cursor-pointer hover:opacity-80 transition-opacity" onclick="openDetailsModalForAdmin(${pData.rawIdx})">
                                    <div class="font-bold text-green-800">${escapeHtml(pData.assignedSub)}</div>
                                    <div class="text-gray-500 mt-1">${escapeHtml(pData.room)}</div>
                                 </td>`;
                    } else {
                        html += `<td class="p-3 text-xs text-center border border-gray-300 bg-yellow-100 align-top cursor-pointer hover:opacity-80 transition-opacity" onclick="openDetailsModalForAdmin(${pData.rawIdx})">
                                    <div class="font-bold text-yellow-800">Unassigned</div>
                                    <div class="text-gray-600 mt-1">${escapeHtml(pData.room)}</div>
                                 </td>`;
                    }
                } else {
                    // Did not request a sub for this period
                    html += `<td class="p-3 text-xs text-center border border-gray-300 bg-gray-200">
                                <div style="width: 100%; height: 2px; background-color: #9ca3af; margin: 10px 0;"></div>
                             </td>`;
                }
            });
            html += `</tr>`;
        });

        tbody.innerHTML = html;
      }

      function renderAdminDataRawSilent(data) {
             adminDataRaw = data || [];
      }

      function loadAdminDataRawSilent() {
        google.script.run
          .withSuccessHandler(renderAdminDataRawSilent)
          .getAdminDashboardData();
      }

      function renderAdminDashboardData(data, renderAtAGlance = false) {
            adminDataRaw = data;
            const today = new Date();
            const offset = today.getTimezoneOffset() * 60000;
            const localISOTime = (new Date(today - offset)).toISOString().split('T')[0];

            if (renderAtAGlance) {
                if (!document.getElementById('atAGlanceDate').value) {
                    document.getElementById('atAGlanceDate').value = localISOTime;
                }
                applyAtAGlanceFilter();
            } else {
                // Default to today
                document.getElementById('filterStartDate').value = localISOTime;
                applyAdminFilters();
            }
          }

      function loadAdminDashboardData(renderAtAGlance = false) {
        document.getElementById('adminDashboardTableBody').innerHTML = '<tr><td colspan="7" class="p-8 text-center text-gray-400 italic">Loading dashboard data...</td></tr>';
        if (renderAtAGlance) {
            document.getElementById('atAGlanceTableBody').innerHTML = '<tr><td class="p-8 text-center text-gray-400 italic">Loading Today at a Glance...</td></tr>';
        }

        google.script.run
          .withSuccessHandler(data => renderAdminDashboardData(data, renderAtAGlance))
          .withFailureHandler(err => {
            document.getElementById('adminDashboardTableBody').innerHTML = `<tr><td colspan="7" class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
            if (renderAtAGlance) {
                document.getElementById('atAGlanceTableBody').innerHTML = `<tr><td class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
            }
          })
          .getAdminDashboardData();
      }

      function applyAdminFilters() {
        const startDateStr = document.getElementById('filterStartDate').value;
        const endDateStr = document.getElementById('filterEndDate').value;
        const searchText = (document.getElementById('filterSearchText').value || '').toLowerCase();

        adminDataFiltered = adminDataRaw.filter(req => {
          let match = true;
          if (startDateStr && req.date < startDateStr) match = false;
          if (endDateStr && req.date > endDateStr) match = false;

          if (match && searchText) {
             const searchString = [
                 req.teacherName,
                 req.assignedSub,
                 req.course,
                 req.date,
                 req.room,
                 req.period
             ].map(val => String(val || '').toLowerCase()).join(' ');

             if (!searchString.includes(searchText)) {
                 match = false;
             }
          }

          return match;
        });

        renderAdminTable();
      }

      function resetAdminFilters() {
        document.getElementById('filterStartDate').value = '';
        document.getElementById('filterEndDate').value = '';
        document.getElementById('filterSearchText').value = '';
        applyAdminFilters();
      }

      function setAdminView(viewType) {
        adminCurrentView = viewType;
        const btnPeriod = document.getElementById('adminTogglePeriod');
        const btnRequest = document.getElementById('adminToggleRequest');

        if (viewType === 'period') {
          btnPeriod.className = "flex-1 py-1.5 px-4 rounded-lg border-2 border-[#00843D] bg-[#00843D] text-white font-bold transition-colors";
          btnRequest.className = "flex-1 py-1.5 px-4 rounded-lg border-2 border-gray-200 text-gray-500 font-bold transition-colors";
        } else {
          btnRequest.className = "flex-1 py-1.5 px-4 rounded-lg border-2 border-[#00843D] bg-[#00843D] text-white font-bold transition-colors";
          btnPeriod.className = "flex-1 py-1.5 px-4 rounded-lg border-2 border-gray-200 text-gray-500 font-bold transition-colors";
        }

        renderAdminTable();
      }

      function sortAdminTable(col) {
        if (adminSortCol === col) {
          adminSortAsc = !adminSortAsc;
        } else {
          adminSortCol = col;
          adminSortAsc = true;
        }
        renderAdminTable();
      }

      function getUniqueRequests(filteredData) {
        const unique = {};
        filteredData.forEach(req => {
          if (!unique[req.id]) {
            unique[req.id] = {
              id: req.id,
              teacherName: req.teacherName,
              teacherEmail: req.teacherEmail,
              date: req.date,
              formDateString: req.formDateString,
              reason: req.reason,
              duration: req.duration,
              periodsString: req.periodsString,
              urgency: req.urgency,
              instructions: req.instructions,
              // Storing one raw index to use if we needed it, though edit modal uses id
              rawIdx: adminDataRaw.findIndex(r => r.id === req.id)
            };
          }
        });
        return Object.values(unique);
      }

      function renderAdminTable() {
        const thead = document.getElementById('adminDashboardTableHeader');
        const tbody = document.getElementById('adminDashboardTableBody');
        if (adminDataFiltered.length === 0) {
          tbody.innerHTML = '<tr><td colspan="7" class="p-8 text-center text-gray-400 italic">No absence requests found for these filters.</td></tr>';
          return;
        }

        if (adminCurrentView === 'period') {
          thead.innerHTML = `
            <tr class="bg-gray-50 border-b border-gray-200 text-sm text-gray-700 uppercase">
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('teacherName')">Name ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('date')">Date ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('period')">Period ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('course')">Course ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('room')">Room ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('assignedSub')">Assigned Sub ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('instructions')">Instructions ↕️</th>
            </tr>
          `;

          // Sort data
          adminDataFiltered.sort((a, b) => {
            let valA = a[adminSortCol] || '';
            let valB = b[adminSortCol] || '';

            if (adminSortCol === 'period') {
               valA = parseInt(valA) || 0;
               valB = parseInt(valB) || 0;
            } else {
               valA = String(valA).toLowerCase();
               valB = String(valB).toLowerCase();
            }

            if (valA < valB) return adminSortAsc ? -1 : 1;
            if (valA > valB) return adminSortAsc ? 1 : -1;
            return 0;
          });

          let html = '';
          adminDataFiltered.forEach((req, idx) => {
            // Find index in raw data for correct modal launching
            const rawIdx = adminDataRaw.findIndex(r => r.id === req.id && r.period === req.period);

            let courseHtml = req.course ? `<span class="inline-block max-w-[150px] sm:max-w-[180px] md:max-w-[240px] truncate align-middle" title="${escapeHtml(req.course)}">${escapeHtml(req.course)}</span>` : '';
            if (req.course === "No Class Assigned" || !req.course) {
               courseHtml = `<span class="tooltip text-red-500 font-bold px-2 py-1 bg-red-50 rounded cursor-help">--<span class="tooltiptext">No course found</span></span>`;
            }

            let roomHtml = escapeHtml(req.room);
            if (req.room === "No Class Assigned" || !req.room) {
               roomHtml = `<span class="tooltip text-red-500 font-bold px-2 py-1 bg-red-50 rounded cursor-help">--<span class="tooltiptext">No room found</span></span>`;
            }

            let subHtml = req.assignedSub ? `<span class="bg-green-100 text-green-800 border border-green-200 px-2 py-1 rounded font-bold">${escapeHtml(req.assignedSub)}</span>` : `<span class="text-gray-400 italic font-medium">Unassigned</span>`;

            html += `
              <tr class="hover:bg-gray-50 transition-colors border-b border-gray-100 last:border-b-0">
                <td class="p-3 text-sm text-gray-800 font-medium"><a href="mailto:${escapeHtml(req.teacherEmail)}" class="hover:underline text-blue-600">${escapeHtml(req.teacherName)}</a></td>
                <td class="p-3 text-sm text-gray-600">${escapeHtml(req.date)}</td>
                <td class="p-3 text-sm font-bold text-gray-700">Period ${escapeHtml(req.period)}</td>
                <td class="p-3 text-sm text-gray-600">${courseHtml}</td>
                <td class="p-3 text-sm text-gray-600">${roomHtml}</td>
                <td class="p-3 text-sm cursor-pointer hover:bg-gray-100" onclick="openDetailsModalForAdmin(${rawIdx})">${subHtml}</td>
                <td class="p-3 text-sm text-gray-500 max-w-xs truncate" title="${escapeHtml(req.instructions) || ''}">${escapeHtml(req.instructions) || '-'}</td>
              </tr>
            `;
          });

          tbody.innerHTML = html;
        } else {
          // Request View
          thead.innerHTML = `
            <tr class="bg-gray-50 border-b border-gray-200 text-sm text-gray-700 uppercase">
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('teacherName')">Name ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('date')">Date ↕️</th>
              <th class="p-3 cursor-pointer hover:bg-gray-100" onclick="sortAdminTable('reason')">Reason ↕️</th>
              <th class="p-3">Duration</th>
              <th class="p-3">Periods</th>
            </tr>
          `;

          let uniqueRequests = getUniqueRequests(adminDataFiltered);

          // Sort unique data
          uniqueRequests.sort((a, b) => {
            let valA = a[adminSortCol] || '';
            let valB = b[adminSortCol] || '';
            valA = String(valA).toLowerCase();
            valB = String(valB).toLowerCase();

            if (valA < valB) return adminSortAsc ? -1 : 1;
            if (valA > valB) return adminSortAsc ? 1 : -1;
            return 0;
          });

          let html = '';
          uniqueRequests.forEach((req, idx) => {
            html += `
              <tr class="hover:bg-gray-100 transition-colors border-b border-gray-100 last:border-b-0 cursor-pointer" onclick="openAdminEditModal('${escapeJs(req.id)}')">
                <td class="p-3 text-sm text-gray-800 font-medium">${escapeHtml(req.teacherName)}</td>
                <td class="p-3 text-sm text-gray-600">${escapeHtml(req.date)}</td>
                <td class="p-3 text-sm text-gray-600">${escapeHtml(req.reason)}</td>
                <td class="p-3 text-sm text-gray-600">${escapeHtml(req.duration)}</td>
                <td class="p-3 text-sm text-gray-600 font-bold">${escapeHtml(req.periodsString)}</td>
              </tr>
            `;
          });

          if (uniqueRequests.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="p-8 text-center text-gray-400 italic">No absence requests found for these filters.</td></tr>';
          } else {
            tbody.innerHTML = html;
          }
        }
      }

      function openDetailsModalForAdmin(index) {
        const req = adminDataRaw[index];
        if (!req) return;

        const assignId = `admin-modal-assign-${req.id}-${req.period}`;

        let html = `
          <div class="bg-gray-50 p-4 rounded-xl border mb-4">
             <div class="grid grid-cols-2 gap-4 text-sm">
                <div><span class="font-bold text-gray-700">Teacher:</span> ${escapeHtml(req.teacherName)}</div>
                <div><span class="font-bold text-gray-700">Email:</span> <a href="mailto:${escapeHtml(req.teacherEmail)}" class="text-blue-600 hover:underline">${escapeHtml(req.teacherEmail)}</a></div>
                <div><span class="font-bold text-gray-700">Date:</span> ${escapeHtml(req.date)}</div>
                <div><span class="font-bold text-gray-700">Period:</span> ${escapeHtml(req.period)}</div>
                <div><span class="font-bold text-gray-700">Room:</span> ${escapeHtml(req.room)}</div>
                <div><span class="font-bold text-gray-700">Course:</span> ${escapeHtml(req.course)}</div>
             </div>
          </div>

          <div class="space-y-3">
             <div>
               <p class="font-bold text-sm text-gray-700">Absence Reason</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.reason) || 'N/A'}</p>
             </div>
             <div>
               <p class="font-bold text-sm text-gray-700">Duration</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.duration) || 'N/A'}</p>
             </div>
             <div>
               <p class="font-bold text-sm text-gray-700">Special Instructions</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.instructions) || 'None provided'}</p>
             </div>
          </div>

          <div class="mt-6 border-t pt-4">
             <p class="font-bold text-sm text-gray-700 mb-2">Assign/Edit Substitute</p>
             <div class="flex items-center gap-2">
                <input type="text" id="${escapeHtml(assignId)}" list="dynamicStaffDataList" onfocus="updateDynamicStaffList('${escapeJs(req.formDateString || req.date)}', '${escapeJs(req.period)}')" value="${escapeHtml(req.assignedSub) || ''}" placeholder="Enter sub name..." class="flex-1 p-3 border rounded-lg text-sm outline-none focus:ring-2 focus:ring-[#00843D] bg-white">
                <button onclick="handleAdminAssign('${escapeJs(req.id)}', '${escapeJs(req.period)}', '${escapeJs(assignId)}')" class="bg-[#00843D] hover:bg-green-700 text-white font-bold py-3 px-6 rounded-lg transition-colors shadow-sm">
                  Save
                </button>
             </div>
             <p class="text-xs text-gray-500 mt-2 italic">To remove a sub, clear the box and click Save.</p>
          </div>
        `;

        document.getElementById('detailsContent').innerHTML = html;

        const modal = document.querySelector('#detailsModal');
        const overlay = document.querySelector('#genericModalOverlay');
        const body = document.querySelector('body');

        modal.classList.remove('opacity-0');
        modal.classList.remove('pointer-events-none');
        overlay.classList.remove('opacity-0');
        overlay.classList.remove('pointer-events-none');
        body.classList.add('modal-active');
      }

      function handleAdminAssign(absenceId, period, selectElementId) {
        const selectEl = document.getElementById(selectElementId);
        let subName = selectEl.value; // It can be empty to unassign!

        // Strip out duty suffix if present: "Name - Duty" -> "Name"
        if (subName) {
           subName = subName.replace(/\s+-\s+.*$/, '');
        }

        if (subName && !staffList.some(s => s.name === subName)) {
          alert("Please enter a valid staff name from the list, or leave it blank to clear.");
          return;
        }

        const btn = selectEl.nextElementSibling;
        const originalText = btn.innerText;
        btn.disabled = true;
        btn.innerText = "Saving...";

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              closeAllModals();

              // Check if At A Glance dashboard is currently visible
              const atAGlanceView = document.getElementById('atAGlanceDashboardView');
              const isAtAGlanceVisible = atAGlanceView && !atAGlanceView.classList.contains('hidden');

              // Refresh admin data behind the scenes
              loadAdminDashboardData(isAtAGlanceVisible);

              // Also refresh quick cover if it's the same 2-day window
              loadQuickCover();
              loadAdminDataRawSilent();
            } else {
              alert("Failed to assign sub: " + (result ? result.error : "Unknown error"));
              btn.disabled = false;
              btn.innerText = originalText;
            }
          })
          .withFailureHandler(err => {
            alert("Error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .assignSubToPeriod(absenceId, period, subName);
      }

      function openDetailsModal(index) {
        const req = currentUnfilledData[index];
        if (!req) return;

        const assignId = `modal-assign-${req.id}-${req.period}`;

        let html = `
          <div class="bg-gray-50 p-4 rounded-xl border mb-4">
             <div class="grid grid-cols-2 gap-4 text-sm">
                <div><span class="font-bold text-gray-700">Teacher:</span> ${escapeHtml(req.teacherName)}</div>
                <div><span class="font-bold text-gray-700">Email:</span> <a href="mailto:${escapeHtml(req.teacherEmail)}" class="text-blue-600 hover:underline">${escapeHtml(req.teacherEmail)}</a></div>
                <div><span class="font-bold text-gray-700">Date:</span> ${escapeHtml(req.date)}</div>
                <div><span class="font-bold text-gray-700">Period:</span> ${escapeHtml(req.period)}</div>
                <div><span class="font-bold text-gray-700">Room:</span> ${escapeHtml(req.room)}</div>
                <div><span class="font-bold text-gray-700">Course:</span> ${escapeHtml(req.course)}</div>
             </div>
          </div>

          <div class="space-y-3">
             <div>
               <p class="font-bold text-sm text-gray-700">Absence Reason</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.reason) || 'N/A'}</p>
             </div>
             <div>
               <p class="font-bold text-sm text-gray-700">Duration</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.duration) || 'N/A'}</p>
             </div>
             <div>
               <p class="font-bold text-sm text-gray-700">Special Instructions</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.instructions) || 'None provided'}</p>
             </div>
          </div>

          <div class="mt-6 border-t pt-4">
             <p class="font-bold text-sm text-gray-700 mb-2">Assign Substitute</p>
             <div class="flex items-center gap-2">
                <input type="text" id="${escapeHtml(assignId)}" list="dynamicStaffDataList" onfocus="updateDynamicStaffList('${escapeJs(req.formDateString || req.date)}', '${escapeJs(req.period)}')" placeholder="Enter sub name..." class="flex-1 p-3 border rounded-lg text-sm outline-none focus:ring-2 focus:ring-[#00843D] bg-white">
                <button onclick="handleAssign('${escapeJs(req.id)}', '${escapeJs(req.period)}', '${escapeJs(assignId)}', true)" class="bg-[#00843D] hover:bg-green-700 text-white font-bold py-3 px-6 rounded-lg transition-colors shadow-sm">
                  Assign
                </button>
             </div>
          </div>
        `;

        document.getElementById('detailsContent').innerHTML = html;

        const body = document.querySelector('body');
        const modal = document.querySelector('#detailsModal');
        const overlay = document.querySelector('#genericModalOverlay');

        modal.classList.remove('opacity-0');
        modal.classList.remove('pointer-events-none');
        overlay.classList.remove('opacity-0');
        overlay.classList.remove('pointer-events-none');
        body.classList.add('modal-active');
      }

      function closeDetailsModal() {
        const body = document.querySelector('body');
        const modal = document.querySelector('#detailsModal');
        const overlay = document.querySelector('#genericModalOverlay');

        modal.classList.add('opacity-0');
        modal.classList.add('pointer-events-none');
        overlay.classList.add('opacity-0');
        overlay.classList.add('pointer-events-none');
        body.classList.remove('modal-active');
      }

      function handleAssign(absenceId, period, selectElementId, isFromModal = false) {
        const selectEl = document.getElementById(selectElementId);
        let subName = selectEl.value;

        if (!subName) {
          alert("Please enter a substitute to assign.");
          return;
        }

        // Strip out duty suffix if present: "Name - Duty" -> "Name"
        subName = subName.replace(/\s+-\s+.*$/, '');

        if (!staffList.some(s => s.name === subName)) {
          alert("Please enter a valid staff name from the list.");
          return;
        }

        const btn = selectEl.nextElementSibling;
        const originalText = btn.innerText;
        btn.disabled = true;
        btn.innerText = "Saving...";

        google.script.run
          .withSuccessHandler(result => {
            if (result && result.success) {
              if (isFromModal) {
                 closeDetailsModal();
              }
              loadQuickCover(); // Refresh the list
              loadAdminDataRawSilent();
            } else {
              alert("Failed to assign sub: " + (result ? result.error : "Unknown error"));
              btn.disabled = false;
              btn.innerText = originalText;
            }
          })
          .withFailureHandler(err => {
            alert("Connection error: " + err.message);
            btn.disabled = false;
            btn.innerText = originalText;
          })
          .assignSubToPeriod(absenceId, period, subName);
      }

      function renderMySubDuties(duties) {
            mySubDutiesData = duties;
            const container = document.getElementById('mySubDuties');

            if (duties.length === 0) {
              container.innerHTML = 'You have no assigned sub-duties over the next week.';
              return;
            }

            let html = '<div class="divide-y divide-gray-200 text-left">';

            duties.forEach((duty, index) => {
              let roomPillClass = duty.room === "No Class Assigned" ? "bg-gray-100 text-gray-600 border-gray-200" : "bg-blue-100 text-blue-800 border-blue-200";
              let coursePillClass = duty.course === "No Class Assigned" ? "bg-gray-100 text-gray-600 border-gray-200" : "bg-purple-100 text-purple-800 border-purple-200";

              html += `
                <div onclick="openSubDutyDetailsModal(${index})" class="p-4 hover:bg-orange-50 transition-colors flex flex-col sm:flex-row justify-between items-start sm:items-center bg-white gap-4 cursor-pointer">
                  <div>
                    <p class="font-bold text-[#002147]">${escapeHtml(duty.date)}</p>
                    <p class="text-sm text-gray-600 font-semibold">${escapeHtml(duty.teacherName)}</p>

                    <div class="mt-2 flex flex-col gap-1 items-start">
                       <span class="inline-block bg-orange-100 text-orange-800 text-xs font-bold px-3 py-1 rounded-full border border-orange-200">Period ${escapeHtml(duty.period)}</span>
                       <span class="inline-block ${roomPillClass} text-xs font-bold px-3 py-1 rounded-full border">Room: ${escapeHtml(duty.room)}</span>
                       <span class="inline-block ${coursePillClass} text-xs font-bold px-3 py-1 rounded-full border truncate max-w-xs">Course: ${escapeHtml(duty.course)}</span>
                    </div>
                  </div>
                  <div class="flex items-center gap-2 w-full sm:w-auto" onclick="event.stopPropagation()">
                    <button onclick="cancelMySubDutyClient('${escapeJs(duty.id)}', '${escapeJs(duty.period)}')" class="bg-white hover:bg-red-50 text-red-600 text-sm font-bold py-2 px-4 rounded-lg border border-red-200 transition-colors shadow-sm">
                      Cancel
                    </button>
                  </div>
                </div>
              `;
            });
            html += '</div>';

            container.innerHTML = html;
            container.classList.remove('text-gray-400');
          }

      function loadMySubDuties() {
        google.script.run
          .withSuccessHandler(renderMySubDuties)
          .withFailureHandler(err => {
            document.getElementById('mySubDuties').innerHTML = `<div class="p-8 text-center text-red-600 font-bold">Failed to load sub-duties: ${err.message}</div>`;
          })
          .getMySubDuties();
      }

      function openSubDutyDetailsModal(index) {
        const req = mySubDutiesData[index];
        if (!req) return;

        let html = `
          <div class="bg-gray-50 p-4 rounded-xl border mb-4">
             <div class="grid grid-cols-2 gap-4 text-sm">
                <div><span class="font-bold text-gray-700">Teacher:</span> ${escapeHtml(req.teacherName)}</div>
                <div><span class="font-bold text-gray-700">Email:</span> <a href="mailto:${escapeHtml(req.teacherEmail)}" class="text-blue-600 hover:underline">${escapeHtml(req.teacherEmail)}</a></div>
                <div><span class="font-bold text-gray-700">Date:</span> ${escapeHtml(req.date)}</div>
                <div><span class="font-bold text-gray-700">Period:</span> ${escapeHtml(req.period)}</div>
                <div><span class="font-bold text-gray-700">Room:</span> ${escapeHtml(req.room)}</div>
                <div><span class="font-bold text-gray-700">Course:</span> ${escapeHtml(req.course)}</div>
             </div>
          </div>

          <div class="space-y-3">
             <div>
               <p class="font-bold text-sm text-gray-700">Special Instructions</p>
               <p class="text-gray-600 bg-white p-3 border rounded-lg">${escapeHtml(req.instructions) || 'None provided'}</p>
             </div>
          </div>
        `;

        document.getElementById('detailsContent').innerHTML = html;

        const body = document.querySelector('body');
        const modal = document.querySelector('#detailsModal');
        const overlay = document.querySelector('#genericModalOverlay');

        modal.classList.remove('opacity-0');
        modal.classList.remove('pointer-events-none');
        overlay.classList.remove('opacity-0');
        overlay.classList.remove('pointer-events-none');
        body.classList.add('modal-active');
      }

      function cancelMySubDutyClient(id, period) {
        if (!confirm("Are you sure you want to cancel your assigned sub coverage for this period? The sub coordinator will be notified.")) return;

        const container = document.getElementById('mySubDuties');
        container.innerHTML = '<div class="spinner mx-auto my-4"></div>';

        google.script.run
          .withSuccessHandler(result => {
             if (result && result.success) {
               alert("Sub-duty cancelled successfully. The sub coordinator has been notified.");
               loadMySubDuties();
               if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 loadQuickCover();
                 loadAdminDataRawSilent();
               }
             } else {
               alert("Failed to cancel sub-duty: " + (result ? result.error : "Unknown"));
               loadMySubDuties();
             }
          })
          .withFailureHandler(err => {
             alert("Error: " + err.message);
             loadMySubDuties();
          })
          .cancelMySubDuty(id, period);
      }

      function renderMyAbsences(absences) {
            myAbsencesData = absences;
            const container = document.getElementById('myAbsences');
            if (absences.length === 0) {
              container.innerHTML = 'No active requests found. You are doing great!';
              return;
            }

            let html = '<div class="space-y-3 text-left">';
            absences.forEach((abs, index) => {
              let bgClass = abs.urgency === 'Urgent' ? 'bg-red-50 border-red-200' : 'bg-gray-50 border-gray-200';
              let textClass = abs.urgency === 'Urgent' ? 'text-red-700' : 'text-[#002147]';

              html += `
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center p-4 border rounded-xl shadow-sm ${bgClass} gap-4">
                  <div>
                    <div class="flex items-center gap-2">
                       <p class="font-bold ${textClass} text-lg">${escapeHtml(abs.date)}</p>
                       <span class="text-[10px] font-bold uppercase tracking-wider opacity-60 bg-white px-2 py-0.5 rounded border border-gray-300">${escapeHtml(abs.urgency)}</span>
                    </div>
                    <p class="text-sm text-gray-600 font-semibold">${escapeHtml(abs.reason)}</p>
                    <p class="text-xs text-gray-500 mt-1">Periods: ${escapeHtml(abs.periods)}</p>
                  </div>
                  <div class="flex gap-2 w-full sm:w-auto">
                    <button onclick="editAbsence(${index})" class="flex-1 sm:flex-none bg-white hover:bg-gray-100 text-gray-700 text-sm font-bold py-2 px-4 rounded-lg border border-gray-300 transition-colors shadow-sm">
                      Edit
                    </button>
                    <button onclick="cancelAbsence('${escapeJs(abs.id)}')" class="flex-1 sm:flex-none bg-white hover:bg-red-50 text-red-600 text-sm font-bold py-2 px-4 rounded-lg border border-red-200 transition-colors shadow-sm">
                      Cancel
                    </button>
                  </div>
                </div>
              `;
            });
            html += '</div>';

            container.innerHTML = html;
            container.classList.remove('text-gray-400');
          }

      function loadMyAbsences() {
        google.script.run
          .withSuccessHandler(renderMyAbsences)
          .withFailureHandler(err => {
            document.getElementById('myAbsences').innerHTML = `<div class="text-red-600 font-bold">Failed to load absences: ${err.message}</div>`;
          })
          .getMyAbsences();
      }

      function editAbsence(index) {
        const abs = myAbsencesData[index];
        if (!abs) return;

        document.getElementById('requestModalTitle').innerText = "Edit Request";
        document.getElementById('editAbsenceId').value = abs.id;
        document.getElementById('submitBtn').innerText = "Save Changes";

        // Date format handling
        if (abs.formDateString) {
           document.getElementById('absDate').value = abs.formDateString;
        } else {
           let d = new Date(abs.rawDateString);
           let year = d.getFullYear();
           let month = String(d.getMonth() + 1).padStart(2, '0');
           let day = String(d.getDate()).padStart(2, '0');
           document.getElementById('absDate').value = `${year}-${month}-${day}`;
        }

        document.getElementById('absDuration').value = abs.duration || "Full Day";

        let pArr = abs.periods.split(",").map(p => p.trim());
        document.querySelectorAll('.period-cb').forEach(cb => {
           cb.checked = pArr.includes(cb.value);
        });

        document.getElementById('absReason').value = abs.reason;
        checkSpecialReason();

        setUrgency(abs.urgency === 'Urgent' ? 'Urgent' : 'Standard');

        // Remove HR prefix if present
        let inst = abs.instructions || "";
        let hrConfirmed = false;
        if (inst.startsWith("[HR Docs Provided] ")) {
           hrConfirmed = true;
           inst = inst.replace("[HR Docs Provided] ", "");
        }
        document.getElementById('specialInst').value = inst;

        if (hrConfirmed && !document.getElementById('hrSection').classList.contains('hidden')) {
           document.getElementById('hrConfirm').checked = true;
        } else {
           document.getElementById('hrConfirm').checked = false;
        }

        document.getElementById('agreement').checked = true; // Pre-check for convenience on edit

        document.getElementById('adminDeleteBtn').classList.add('hidden'); // Regular users don't see delete in the modal, they use the cancel button on the list
        toggleModal(true);
        setTimeout(captureModalState, 50);
      }

      function openAdminEditModal(reqId) {
         // Find the request from adminDataFiltered based on id
         const req = adminDataFiltered.find(r => r.id === reqId);
         if (!req) return;

         document.getElementById('requestModalTitle').innerText = "Edit Request (Admin)";
         document.getElementById('editAbsenceId').value = req.id;
         document.getElementById('submitBtn').innerText = "Save Changes";
         document.getElementById('adminDeleteBtn').classList.remove('hidden');

         // Handle date
         if (req.formDateString) {
             document.getElementById('absDate').value = req.formDateString;
         } else {
             document.getElementById('absDate').value = req.date;
         }

         document.getElementById('absDuration').value = req.duration || "Full Day";

         let pArr = (req.periodsString || "").split(",").map(p => p.trim());
         document.querySelectorAll('.period-cb').forEach(cb => {
            cb.checked = pArr.includes(cb.value);
         });

         document.getElementById('absReason').value = req.reason;
         checkSpecialReason();

         setUrgency(req.urgency.includes('Urgent') ? 'Urgent' : 'Standard');

         // Parse instructions for HR confirmation flag
         let inst = req.instructions || "";
         let hrConfirmed = false;
         if (inst.startsWith("[HR Docs Provided] ")) {
            hrConfirmed = true;
            inst = inst.replace("[HR Docs Provided] ", "");
         }
         document.getElementById('specialInst').value = inst;

         if (hrConfirmed && !document.getElementById('hrSection').classList.contains('hidden')) {
            document.getElementById('hrConfirm').checked = true;
         } else {
            document.getElementById('hrConfirm').checked = false;
         }

         document.getElementById('agreement').checked = true;

         toggleModal(true);
         setTimeout(captureModalState, 50);
      }

      function deleteAbsenceFromModal() {
         const id = document.getElementById('editAbsenceId').value;
         if (!id) return;

         if (!confirm("Are you sure you want to completely delete this absence request? The teacher will be notified.")) return;

         const btn = document.getElementById('adminDeleteBtn');
         const originalText = btn.innerText;
         btn.disabled = true;
         btn.innerText = "Deleting...";

         google.script.run
           .withSuccessHandler(result => {
              if (result && result.success) {
                alert("Request cancelled successfully.");
                toggleModal(); // close modal
                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
                  // Admin dashboard view handles its own UI refresh when raw data is reloaded, but let's force a refresh
                  loadAdminDashboardData();
                }
              } else {
                alert("Failed to cancel: " + (result ? result.error : "Unknown"));
                btn.disabled = false;
                btn.innerText = originalText;
              }
           })
           .withFailureHandler(err => {
              alert("Error: " + err.message);
              btn.disabled = false;
              btn.innerText = originalText;
           })
           .cancelAbsence(id);
      }

      function cancelAbsence(id) {
        if (!confirm("Are you sure you want to cancel this request? Any assigned substitutes will be automatically notified.")) return;

        const container = document.getElementById('myAbsences');
        container.innerHTML = '<div class="spinner mx-auto my-4"></div>';

        google.script.run
          .withSuccessHandler(result => {
             if (result && result.success) {
               alert("Request cancelled successfully.");
               loadMyAbsences();
               if (!document.getElementById('adminSection').classList.contains('hidden')) {
                 loadQuickCover();
                 loadAdminDataRawSilent();
               }
             } else {
               alert("Failed to cancel: " + (result ? result.error : "Unknown"));
               loadMyAbsences();
             }
          })
          .withFailureHandler(err => {
             alert("Error: " + err.message);
             loadMyAbsences();
          })
          .cancelAbsence(id);
      }

      document.getElementById('absenceForm').onsubmit = function(e) {
        e.preventDefault();

        const selectedPeriods = Array.from(document.querySelectorAll('.period-cb:checked'))
                                     .map(cb => cb.value).join(', ');

        if (!selectedPeriods) {
          alert("Please select at least one period for coverage.");
          return;
        }

        const reasonSelect = document.getElementById('absReason');
        const reason = reasonSelect.value;
        const hrChecked = document.getElementById('hrConfirm').checked;
        const agreed = document.getElementById('agreement').checked;

        let hrRequired = false;
        if (reasonSelect.options.length > 0 && reasonSelect.selectedIndex > -1) {
            hrRequired = reasonSelect.options[reasonSelect.selectedIndex].getAttribute('data-hr-required') === 'true';
        }

        if (hrRequired && !hrChecked) {
          alert("Please confirm HR documentation has been provided.");
          return;
        }
        if (!agreed) {
          alert("You must agree to the terms to submit.");
          return;
        }

        const btn = document.getElementById('submitBtn');
        btn.disabled = true;
        btn.innerHTML = '<div class="flex items-center justify-center gap-2"><div class="spinner"></div> Sending...</div>';

        const formData = {
          date: document.getElementById('absDate').value,
          duration: document.getElementById('absDuration').value,
          periods: selectedPeriods,
          reason: reason,
          urgency: document.getElementById('urgencyValue').value,
          specialInstructions: document.getElementById('specialInst').value,
          hrConfirmed: hrChecked
        };

        const editId = document.getElementById('editAbsenceId').value;

        if (editId) {
          google.script.run
            .withSuccessHandler((result) => {
              if (result && result.success) {
                alert("Request Updated Successfully!");
                btn.disabled = false;
                btn.innerHTML = 'Save Changes';
                toggleModal();
                loadMyAbsences();
                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
                  // Admin dashboard view handles its own UI refresh when raw data is reloaded, but let's force a refresh
                  loadAdminDashboardData();
                }
              } else {
                alert("Database Error: " + (result ? result.error : "Unknown issue"));
                btn.disabled = false;
                btn.innerHTML = 'Save Changes';
              }
            })
            .withFailureHandler(err => {
              alert("Connection Error: " + err.message);
              btn.disabled = false;
              btn.innerHTML = 'Save Changes';
            })
            .updateAbsence(editId, formData);
        } else {
          google.script.run
            .withSuccessHandler((result) => {
              if (result && result.success) {
                alert("Request Submitted Successfully!");
                document.getElementById('absenceForm').reset();
                setUrgency('Standard');
                document.getElementById('hrSection').classList.add('hidden');
                btn.disabled = false;
                btn.innerHTML = 'Submit Request';
                toggleModal();
                loadMyAbsences();

                if (!document.getElementById('adminSection').classList.contains('hidden')) {
                  loadQuickCover();
                  loadAdminDataRawSilent();
                }
              } else {
                alert("Database Error: " + (result ? result.error : "Unknown issue"));
                btn.disabled = false;
                btn.innerHTML = 'Submit Request';
              }
            })
            .withFailureHandler(err => {
              alert("Connection Error: " + err.message);
              btn.disabled = false;
              btn.innerHTML = 'Submit Request';
            })
            .submitAbsence(formData);
        }
      };

      function toggleHRDashboard() {
        const mainContent = document.getElementById('mainContent');
        const adminView = document.getElementById('adminDashboardView');
        const hrView = document.getElementById('hrDashboardView');
        const atAGlanceView = document.getElementById('atAGlanceDashboardView');
        const settingsView = document.getElementById('settingsDashboardView');

        if (hrView.classList.contains('hidden')) {
          mainContent.classList.add('hidden');
          if (adminView) adminView.classList.add('hidden');
          if (atAGlanceView) atAGlanceView.classList.add('hidden');
          if (settingsView) settingsView.classList.add('hidden');
          hrView.classList.remove('hidden');

          if (hrDataRaw.length === 0) {
            loadHRDashboardData();
          }
        } else {
          hrView.classList.add('hidden');
          mainContent.classList.remove('hidden');
        }
      }

      function setHRView(viewType) {
        hrCurrentView = viewType;
        const btnAbsences = document.getElementById('hrToggleAbsences');
        const btnSubs = document.getElementById('hrToggleSubs');

        if (viewType === 'absences') {
          btnAbsences.className = "flex-1 py-2 px-4 rounded-lg border-2 border-[#00843D] bg-[#00843D] text-white font-bold transition-colors";
          btnSubs.className = "flex-1 py-2 px-4 rounded-lg border-2 border-gray-200 text-gray-500 font-bold transition-colors";
        } else {
          btnSubs.className = "flex-1 py-2 px-4 rounded-lg border-2 border-[#00843D] bg-[#00843D] text-white font-bold transition-colors";
          btnAbsences.className = "flex-1 py-2 px-4 rounded-lg border-2 border-gray-200 text-gray-500 font-bold transition-colors";
        }

        applyHRFilters();
      }

      function renderHRDashboardData(data) {
            hrDataRaw = data.requests || [];
            hrPayPeriods = data.payPeriods || [];

            const today = new Date();
            const offset = today.getTimezoneOffset() * 60000;
            const endLocalDate = new Date(today - offset);
            const todayStr = endLocalDate.toISOString().split('T')[0];

            if (hrPayPeriods && hrPayPeriods.length > 0) {
              const ppContainer = document.getElementById('hrPayPeriodContainer');
              const ppSelect = document.getElementById('hrPayPeriodSelect');
              ppContainer.classList.remove('hidden');

              let html = `<option value="custom">Custom Range</option>`;
              let selectedIdx = -1;

              // Find the most recently completed pay period
              // "Completed" means today's date is strictly AFTER the pay period's end date
              let mostRecentCompletedIndex = -1;
              let mostRecentEndDate = "";

              for (let i = 0; i < hrPayPeriods.length; i++) {
                 const pp = hrPayPeriods[i];
                 if (todayStr > pp.endDate) {
                    if (mostRecentEndDate === "" || pp.endDate > mostRecentEndDate) {
                       mostRecentEndDate = pp.endDate;
                       mostRecentCompletedIndex = i;
                    }
                 }
              }

              hrPayPeriods.forEach((pp, idx) => {
                 let [sYear, sMonth, sDay] = pp.startDate.split('-');
                 let [eYear, eMonth, eDay] = pp.endDate.split('-');
                 const displayStart = `${sMonth}/${sDay}/${sYear}`;
                 const displayEnd = `${eMonth}/${eDay}/${eYear}`;
                 const isSelected = (idx === mostRecentCompletedIndex) ? "selected" : "";
                 if (isSelected) selectedIdx = idx;

                 html += `<option value="${idx}" ${isSelected}>Pay Period ${pp.periodNumber} (${displayStart} - ${displayEnd})</option>`;
              });

              ppSelect.innerHTML = html;

              if (selectedIdx !== -1) {
                 document.getElementById('hrFilterStartDate').value = hrPayPeriods[selectedIdx].startDate;
                 document.getElementById('hrFilterEndDate').value = hrPayPeriods[selectedIdx].endDate;
              } else {
                 // Fallback if none are completed yet
                 const startLocalDate = new Date(endLocalDate);
                 startLocalDate.setDate(startLocalDate.getDate() - 14);
                 document.getElementById('hrFilterStartDate').value = startLocalDate.toISOString().split('T')[0];
                 document.getElementById('hrFilterEndDate').value = todayStr;
                 ppSelect.value = "custom";
              }

            } else {
              // Default fallback behavior if no pay periods exist
              const startLocalDate = new Date(endLocalDate);
              startLocalDate.setDate(startLocalDate.getDate() - 14);
              const startDateStr = startLocalDate.toISOString().split('T')[0];

              document.getElementById('hrFilterStartDate').value = startDateStr;
              document.getElementById('hrFilterEndDate').value = todayStr;
              document.getElementById('hrPayPeriodContainer').classList.add('hidden');
            }

            applyHRFilters();
          }

      function loadHRDashboardData() {
        document.getElementById('hrDashboardTableBody').innerHTML = '<tr><td colspan="2" class="p-8 text-center text-gray-400 italic">Loading HR dashboard data...</td></tr>';

        google.script.run
          .withSuccessHandler(renderHRDashboardData)
          .withFailureHandler(err => {
            document.getElementById('hrDashboardTableBody').innerHTML = `<tr><td colspan="2" class="p-8 text-center text-red-600 font-bold">Failed to load: ${err.message}</td></tr>`;
          })
          .getHRDashboardData();
      }

      function handleHRDateChange() {
         const ppSelect = document.getElementById('hrPayPeriodSelect');
         if (ppSelect && !document.getElementById('hrPayPeriodContainer').classList.contains('hidden')) {
             const startDateStr = document.getElementById('hrFilterStartDate').value;
             const endDateStr = document.getElementById('hrFilterEndDate').value;

             // Check if dates match any predefined pay period
             let matchFound = false;
             for (let i = 0; i < hrPayPeriods.length; i++) {
                 if (hrPayPeriods[i].startDate === startDateStr && hrPayPeriods[i].endDate === endDateStr) {
                     ppSelect.value = i.toString();
                     matchFound = true;
                     break;
                 }
             }

             if (!matchFound) {
                 ppSelect.value = "custom";
             }
         }
         applyHRFilters();
      }

      function applyPayPeriodSelection() {
         const ppSelect = document.getElementById('hrPayPeriodSelect');
         const val = ppSelect.value;
         if (val !== "custom" && hrPayPeriods[val]) {
             document.getElementById('hrFilterStartDate').value = hrPayPeriods[val].startDate;
             document.getElementById('hrFilterEndDate').value = hrPayPeriods[val].endDate;
         }
         applyHRFilters();
      }

      function applyHRFilters() {
        const startDateStr = document.getElementById('hrFilterStartDate').value;
        const endDateStr = document.getElementById('hrFilterEndDate').value;

        // Filter data by date range
        let filteredData = hrDataRaw.filter(req => {
          let match = true;
          if (startDateStr && req.date < startDateStr) match = false;
          if (endDateStr && req.date > endDateStr) match = false;
          return match;
        });

        let aggregated = {};

        // Extract expected reasons from the DOM dropdown
        let reasonOptions = Array.from(document.getElementById('absReason').options);
        let expectedReasons = reasonOptions.map(opt => opt.value);
        let expectedLabels = reasonOptions.map(opt => opt.text);

        let dynamicReasons = [...expectedReasons];

        if (hrCurrentView === 'absences') {
            filteredData.forEach(req => {
                let teacher = req.teacherName;
                if (!aggregated[teacher]) {
                    aggregated[teacher] = { total: 0 };
                    // Initialize all expected reasons
                    expectedReasons.forEach(r => aggregated[teacher][r] = 0);
                }
                let hours = (req.duration === 'Half Day' ? 4 : 8);

                let reason = req.reason ? req.reason.trim() : "Other";
                let matchedReason = null;

                // Case-insensitive matching against values and labels
                for (let i = 0; i < expectedReasons.length; i++) {
                    if (reason.toLowerCase() === expectedReasons[i].toLowerCase() ||
                        reason.toLowerCase() === expectedLabels[i].toLowerCase()) {
                        matchedReason = expectedReasons[i];
                        break;
                    }
                }

                if (matchedReason) {
                    aggregated[teacher][matchedReason] += hours;
                } else {
                    // It's a legacy or unknown reason
                    let existingDynamic = dynamicReasons.find(r => r.toLowerCase() === reason.toLowerCase());
                    if (existingDynamic) {
                        matchedReason = existingDynamic;
                    } else {
                        matchedReason = reason;
                        dynamicReasons.push(matchedReason);
                    }
                    if (aggregated[teacher][matchedReason] === undefined) {
                        aggregated[teacher][matchedReason] = 0;
                    }
                    aggregated[teacher][matchedReason] += hours;
                }

                aggregated[teacher].total += hours;
            });
        } else {
            // Precompute duty lookup map
            let dutyLookup = {};
            if (staffList && staffList.length > 0) {
                staffList.forEach(s => {
                    dutyLookup[s.name.trim()] = s.duty ? String(s.duty).trim() : null;
                });
            }

            filteredData.forEach(req => {
                req.assignedSubs.forEach(subObj => {
                    let subName = subObj.name;
                    let coveredPeriod = subObj.period;

                    if (!aggregated[subName]) {
                        aggregated[subName] = { duties: 0, extras: 0, total: 0 };
                    }

                    let teacherDuty = dutyLookup[subName];
                    if (teacherDuty && teacherDuty === coveredPeriod) {
                        aggregated[subName].duties += 1;
                    } else {
                        aggregated[subName].extras += 1;
                    }
                    aggregated[subName].total += 1;
                });
            });
        }

        // Convert to array and filter out 0s
        let resultList = [];
        for (let name in aggregated) {
            if (hrCurrentView === 'absences') {
                if (aggregated[name].total > 0) {
                    resultList.push({ name: name, ...aggregated[name] });
                }
            } else {
                if (aggregated[name].total > 0) {
                    resultList.push({ name: name, duties: aggregated[name].duties, extras: aggregated[name].extras, total: aggregated[name].total });
                }
            }
        }

        // Sort alphabetically
        resultList.sort((a, b) => a.name.localeCompare(b.name));

        renderHRTable(resultList, dynamicReasons, expectedReasons, expectedLabels);
      }

      function renderHRTable(dataList, dynamicReasons, expectedReasons, expectedLabels) {
        const thead = document.getElementById('hrDashboardTableHeader');
        const tbody = document.getElementById('hrDashboardTableBody');

        // Dynamically set headers based on view
        if (hrCurrentView === 'absences') {
          let headerHtml = `
            <tr class="bg-gray-50 border-b border-gray-200 text-sm text-gray-700 uppercase">
              <th class="p-3">Teacher Name</th>
          `;

          dynamicReasons.forEach(reason => {
             let label = reason;
             let idx = expectedReasons.findIndex(r => r === reason);
             if (idx !== -1) {
                 label = expectedLabels[idx];
             }
             if (label === "Professional Development") {
                 label = "PD";
             }
             headerHtml += `<th class="p-3">${label}</th>`;
          });

          headerHtml += `<th class="p-3 text-right">Total Hours</th></tr>`;
          thead.innerHTML = headerHtml;
        } else {
          thead.innerHTML = `
            <tr class="bg-gray-50 border-b border-gray-200 text-sm text-gray-700 uppercase">
              <th class="p-3">Teacher Name</th>
              <th class="p-3 text-right">Duties</th>
              <th class="p-3 text-right">Extras</th>
              <th class="p-3 text-right">Totals</th>
            </tr>
          `;
        }

        const colCount = hrCurrentView === 'absences' ? (2 + (dynamicReasons ? dynamicReasons.length : 0)) : 4;

        if (dataList.length === 0) {
          tbody.innerHTML = `<tr><td colspan="${colCount}" class="p-8 text-center text-gray-400 italic">No data found for the selected dates.</td></tr>`;
          return;
        }

        let html = '';
        dataList.forEach(item => {
          if (hrCurrentView === 'absences') {
            html += `
              <tr class="hover:bg-gray-50 transition-colors border-b border-gray-100 last:border-b-0">
                <td class="p-3 text-sm text-gray-800 font-medium">${escapeHtml(item.name)}</td>
            `;

            dynamicReasons.forEach(reason => {
                html += `<td class="p-3 text-sm text-gray-600 font-bold">${escapeHtml(item[reason]) || 0}</td>`;
            });

            html += `
                <td class="p-3 text-sm text-gray-600 font-bold text-right">${escapeHtml(item.total) || 0}</td>
              </tr>
            `;
          } else {
            html += `
              <tr class="hover:bg-gray-50 transition-colors border-b border-gray-100 last:border-b-0">
                <td class="p-3 text-sm text-gray-800 font-medium">${escapeHtml(item.name)}</td>
                <td class="p-3 text-sm text-gray-600 font-bold text-right">${escapeHtml(item.duties)}</td>
                <td class="p-3 text-sm text-gray-600 font-bold text-right">${escapeHtml(item.extras)}</td>
                <td class="p-3 text-sm text-gray-600 font-bold text-right">${escapeHtml(item.total)}</td>
              </tr>
            `;
          }
        });

        tbody.innerHTML = html;
      }

      function copyHRData() {
        const table = document.querySelector('#hrDashboardView table');
        if (!table) return;

        let tsv = '';

        // Get headers
        const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.innerText);
        tsv += headers.join('\t') + '\n';

        // Get rows
        const rows = Array.from(table.querySelectorAll('tbody tr'));
        rows.forEach(tr => {
            // Check if this is the "Loading" or "No data" row
            const cols = tr.querySelectorAll('td');
            if (cols.length === 1 && cols[0].hasAttribute('colspan')) return;

            const rowData = Array.from(cols).map(td => td.innerText);
            tsv += rowData.join('\t') + '\n';
        });

        if (tsv.trim() === '') {
            alert('No data to copy.');
            return;
        }

        navigator.clipboard.writeText(tsv).then(() => {
            const btn = document.querySelector('button[onclick="copyHRData()"]');
            const originalHTML = btn.innerHTML;
            btn.innerHTML = '<span>✅</span> Copied!';
            setTimeout(() => {
                btn.innerHTML = originalHTML;
            }, 2000);
        }).catch(err => {
            console.error('Failed to copy text: ', err);
            alert('Failed to copy data to clipboard.');
        });
      }

      function copyAtAGlance() {
        const container = document.getElementById('atAGlanceCopyContainer');
        const footer = document.getElementById('atAGlanceFooter');

        // Show the footer temporarily for the copy
        footer.classList.remove('hidden');

        // Create a range and selection to copy HTML content
        const range = document.createRange();
        range.selectNode(container);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        try {
            document.execCommand('copy');

            const btn = document.querySelector('button[onclick="copyAtAGlance()"]');
            const originalHTML = btn.innerHTML;
            btn.innerHTML = '<span>✅</span> Copied!';
            setTimeout(() => {
                btn.innerHTML = originalHTML;
            }, 2000);
        } catch (err) {
            console.error('Failed to copy text: ', err);
            alert('Failed to copy data to clipboard.');
        }

        // Clean up
        selection.removeAllRanges();
        footer.classList.add('hidden');
      }
