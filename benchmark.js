const fs = require('fs');
const { performance } = require('perf_hooks');

// Mocks
let getSS = () => ({
  getSheetByName: () => mockSheet
});
let getUserData = () => ({});
let assertPermission = () => {};
let getSheetOrThrow = () => mockSheet;
let logAuditAction = () => {};
const CacheService = {
  getScriptCache: () => ({ remove: () => {} })
};
let _globalSettingsCache = null;

let mockData = [
  ['Key', 'Value'],
  ['Setting1', 'Val1'],
  ['Setting2', 'Val2'],
  ['Setting3', 'Val3'],
  ['Setting4', 'Val4'],
  ['Setting5', 'Val5'],
  ['Setting6', 'Val6'],
  ['Setting7', 'Val7'],
  ['Setting8', 'Val8'],
  ['Setting9', 'Val9'],
  ['Setting10', 'Val10'],
];

let setValueCount = 0;
let setValuesCount = 0;

let mockSheet = {
  getDataRange: () => ({
    getValues: () => JSON.parse(JSON.stringify(mockData))
  }),
  getRange: (row, col, numRows, numCols) => {
    return {
      setValue: (val) => {
        setValueCount++;
        // Simulate delay
        let start = performance.now();
        while(performance.now() - start < 10) {}
      },
      setValues: (vals) => {
        setValuesCount++;
        // Simulate delay for setValues
        let start = performance.now();
        while(performance.now() - start < 20) {}
      }
    };
  },
  getLastRow: () => mockData.length
};

const newSettings = {
  'Setting1': 'NewVal1',
  'Setting2': 'NewVal2',
  'Setting3': 'NewVal3',
  'Setting4': 'NewVal4',
  'Setting5': 'NewVal5',
  'Setting6': 'NewVal6',
  'Setting7': 'NewVal7',
  'Setting8': 'NewVal8',
  'Setting9': 'NewVal9',
  'Setting10': 'NewVal10',
  'NewSetting1': 'NewVal11',
};

// 1. Original code simulation
function originalUpdateSettings(newSettings) {
  var data = mockSheet.getDataRange().getValues();
  var settingsMap = {};
  for (var i = 1; i < data.length; i++) {
    settingsMap[String(data[i][0]).trim()] = i + 1;
  }
  var rowsToAppend = [];
  var updates = [];
  for (var key in newSettings) {
    if (settingsMap[key]) {
      updates.push({ row: settingsMap[key], val: newSettings[key] });
    } else {
      rowsToAppend.push([key, newSettings[key]]);
    }
  }
  for (var u = 0; u < updates.length; u++) {
    mockSheet.getRange(updates[u].row, 2).setValue(updates[u].val);
  }
  if (rowsToAppend.length > 0) {
    mockSheet.getRange(mockSheet.getLastRow() + 1, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
  }
}

// 2. Optimized code simulation
function optimizedUpdateSettings(newSettings) {
  var data = mockSheet.getDataRange().getValues();
  var settingsMap = {};
  for (var i = 1; i < data.length; i++) {
    settingsMap[String(data[i][0]).trim()] = i;
  }
  var rowsToAppend = [];
  var dataChanged = false;
  for (var key in newSettings) {
    if (settingsMap[key] !== undefined) {
      if (data[settingsMap[key]][1] !== newSettings[key]) {
        data[settingsMap[key]][1] = newSettings[key];
        dataChanged = true;
      }
    } else {
      rowsToAppend.push([key, newSettings[key]]);
    }
  }
  if (dataChanged) {
    mockSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
  if (rowsToAppend.length > 0) {
    mockSheet.getRange(mockSheet.getLastRow() + 1, 1, rowsToAppend.length, 2).setValues(rowsToAppend);
  }
}

console.log("Running baseline (original)...");
setValueCount = 0;
setValuesCount = 0;
let t0 = performance.now();
originalUpdateSettings(newSettings);
let t1 = performance.now();
console.log(`Original Time: ${(t1 - t0).toFixed(2)} ms (setValue calls: ${setValueCount}, setValues calls: ${setValuesCount})`);

console.log("Running optimized...");
setValueCount = 0;
setValuesCount = 0;
t0 = performance.now();
optimizedUpdateSettings(newSettings);
t1 = performance.now();
console.log(`Optimized Time: ${(t1 - t0).toFixed(2)} ms (setValue calls: ${setValueCount}, setValues calls: ${setValuesCount})`);

