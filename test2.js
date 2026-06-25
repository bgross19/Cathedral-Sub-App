      function hasUnsavedData() {
        const isAddRoleModalOpen = !document.getElementById('addRoleModal').classList.contains('opacity-0');
        const isRequestModalOpen = !document.querySelector('#modal').classList.contains('opacity-0');

        if (isAddRoleModalOpen) {
          const email = document.getElementById('addRoleEmail').value;
          if (email && email.trim() !== "") return true;
        }

        if (isRequestModalOpen) {
          const date = document.getElementById('absDate').value;
          const specialInst = document.getElementById('specialInst').value;
          const periodsChecked = document.querySelectorAll('.period-cb:checked').length > 0;

          if (date || (specialInst && specialInst.trim() !== "") || periodsChecked) return true;
        }

        return false;
      }

      function attemptCloseAllModals() {
        if (hasUnsavedData()) {
          if (!confirm("You have unsaved changes. Are you sure you want to close this modal and discard them?")) {
            return;
          }
        }
        closeAllModals();
      }
