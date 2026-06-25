const fs = require('fs');

let html = fs.readFileSync('Index.html', 'utf8');

const jsToAdd = `
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
`;

html = html.replace('function closeAllModals() {', jsToAdd + '\n      function closeAllModals() {');
fs.writeFileSync('Index.html', html);
