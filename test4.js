let originalModalData = {};

// When opening edit modals, we will capture the state.
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
  }
}

function hasUnsavedData() {
  const isAddRoleModalOpen = !document.getElementById('addRoleModal').classList.contains('opacity-0');
  const isRequestModalOpen = !document.querySelector('#modal').classList.contains('opacity-0');

  if (isAddRoleModalOpen) {
    const currentEmail = document.getElementById('addRoleEmail').value;
    const currentRole = document.getElementById('addRoleSelect').value;

    // In add mode, any typed email is unsaved data.
    if (originalModalData.email === undefined && currentEmail.trim() !== "") return true;

    // In edit mode (if we ever use it here), if it changed.
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
      if (currentDate !== originalModalData.date ||
          currentDuration !== originalModalData.duration ||
          currentReason !== originalModalData.reason ||
          currentUrgency !== originalModalData.urgency ||
          currentSpecialInst !== originalModalData.specialInst ||
          currentPeriods !== originalModalData.periods) {
          return true;
      }
    }
  }

  return false;
}
