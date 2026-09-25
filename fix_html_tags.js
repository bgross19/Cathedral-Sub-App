const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

const toReplace = `
      function closeDuplicateWarning() {
        const modal = document.getElementById('duplicateAbsenceModal');
        modal.classList.add('opacity-0', 'pointer-events-none');
      }

      function cancelDuplicateWarning() {
        closeDuplicateWarning();
        toggleModal(); // Close the new request form
      }

      // Override attemptCloseAllModals to use showConfirm
      let oldAttemptCloseAllModals = "";
      try {
        oldAttemptCloseAllModals = attemptCloseAllModals.toString();
      } catch(e) {
        console.warn("SecurityError caught: ", e);
      }
      // Wait, we need to inject this properly. Let's do it via regex on the actual function body later.
    </script>
`;

const replacement = `
      function closeDuplicateWarning() {
        const modal = document.getElementById('duplicateAbsenceModal');
        modal.classList.add('opacity-0', 'pointer-events-none');
      }

      function cancelDuplicateWarning() {
        closeDuplicateWarning();
        toggleModal(); // Close the new request form
      }

      // Override attemptCloseAllModals to use showConfirm
      let oldAttemptCloseAllModals = "";
      try {
        oldAttemptCloseAllModals = attemptCloseAllModals.toString();

        // Let's actually override it properly
        window.attemptCloseAllModals = function() {
          if (hasUnsavedData()) {
            showConfirm("You have unsaved changes. Are you sure you want to close this modal and discard them?", () => {
               originalModalData = {};
               closeAllModals();
            });
            return;
          }

          // Reset the captured state
          originalModalData = {};
          closeAllModals();
        };
      } catch(e) {
        console.warn("Error overriding attemptCloseAllModals: ", e);
      }
    </script>
`;

if (html.includes(toReplace)) {
  html = html.replace(toReplace, replacement);
  fs.writeFileSync('Index.html', html);
  console.log("Replaced successfully!");
} else {
  console.log("Could not find the target string!");
}
