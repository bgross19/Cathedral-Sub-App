const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

const brokenString = `
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

console.log(html.includes(brokenString));
