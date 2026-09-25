const fs = require('fs');
let html = fs.readFileSync('Index.html', 'utf8');

// The original PR added this exact string:
const brokenString = `      // Override attemptCloseAllModals to use showConfirm
      let oldAttemptCloseAllModals = "";
      try {
        oldAttemptCloseAllModals = attemptCloseAllModals.toString();
      } catch(e) {
        console.warn("SecurityError caught: ", e);
      }
      // Wait, we need to inject this properly. Let's do it via regex on the actual function body later.`;

// What we want:
const correctString = `      // Override attemptCloseAllModals to use showConfirm
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
      }`;

// Currently, it looks like it's ALREADY fixed in my checkout? Wait, looking at lines 6950-7050 from `git checkout Index.html` output, it was restored!
