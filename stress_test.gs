function runStressTest() {
  Logger.log("Starting stress test...");

  // Mock formData
  var formData = {
    date: '2025-05-10',
    periods: '1, 2, 3',
    reason: 'Personal',
    duration: 'Full Day',
    urgency: 'Standard',
    specialInstructions: 'Test stress',
    hrConfirmed: false
  };

  // Warning: running stress tests from GAS directly runs into quota issues for emails
  // For demonstration, let's just attempt 5 submissions concurrently.
  // Note: Apps Script doesn't have true multithreading, so simulating concurrency is tricky.
  // We can write a local script using the runner that rapidly calls the functions.
}
