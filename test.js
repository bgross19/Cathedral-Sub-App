function checkUnsavedData() {
  const isAddRoleModalOpen = !document.getElementById('addRoleModal').classList.contains('opacity-0');
  const isRequestModalOpen = !document.querySelector('#modal').classList.contains('opacity-0');

  if (isAddRoleModalOpen) {
    const email = document.getElementById('addRoleEmail').value;
    if (email) return true;
  }

  if (isRequestModalOpen) {
    // Only check if it's not edit mode (add mode has empty editAbsenceId)
    // Actually, editing could also have unsaved data.
    const isEdit = document.getElementById('editAbsenceId').value !== "";
    const date = document.getElementById('absDate').value;
    const notes = document.getElementById('absNotes').value;
    const periodsChecked = document.querySelectorAll('.period-cb:checked').length > 0;

    // For simplicity, just checking if ANY form field has some content that would be lost
    // In edit mode, we could compare to original, but let's just confirm if they click outside.
    if (date || notes || periodsChecked) return true;
  }

  return false;
}
