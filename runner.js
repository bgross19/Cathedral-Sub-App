const fs = require('fs');
const files = ['code.gs', 'powerschool.gs', 'tests.gs'];
let code = '';

// Global mocks
global.Logger = { log: console.log };
global.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (key) => null,
    setProperty: (key, value) => null
  })
};
global.CacheService = {
  getScriptCache: () => ({
    get: (key) => null,
    put: (key, value, exp) => null,
    remove: (key) => null
  })
};
global.UrlFetchApp = {
  fetch: (url, options) => ({
    getResponseCode: () => 200,
    getContentText: () => '{}'
  })
};

global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => ({
      getDataRange: () => ({ getValues: () => [] }),
      appendRow: (row) => null,
      clearContents: () => null,
      getRange: () => ({ setValues: () => null }),
      deleteRows: (startRow, numRows) => null
    })
  })
};

global.Session = {
  getActiveUser: () => ({
    getEmail: () => 'test@example.com'
  }),
  getScriptTimeZone: () => 'America/New_York'
};

global.Utilities = {
  formatDate: (date, tz, format) => date.toISOString(),
  getUuid: () => 'test-uuid'
};

global.LockService = {
  getScriptLock: () => ({
    tryLock: (time) => true,
    waitLock: (time) => null,
    releaseLock: () => null
  })
};

for (const file of files) {
  if (fs.existsSync(file)) {
    code += fs.readFileSync(file, 'utf8') + '\n';
  }
}

try {
  eval(code);
  if (typeof runTests === 'function') {
    runTests();
  } else {
    console.log('No runTests function found.');
  }
} catch (e) {
  console.error(e);
  process.exit(1);
}
