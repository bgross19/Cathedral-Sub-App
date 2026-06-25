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
