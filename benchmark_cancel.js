const { performance } = require('perf_hooks');

// Generate mock data
let mockData = [];
for (let i = 0; i < 1000; i++) {
  let row = ['ID' + i];
  for (let j = 1; j < 20; j++) {
    row.push('Col' + j + 'Val' + i);
  }
  mockData.push(row);
}

let mockSheet = {
  getDataRange: () => ({
    getValues: () => JSON.parse(JSON.stringify(mockData))
  }),
  getRange: (row, col) => {
    return {
      getValue: () => {
        // simulate API delay
        let start = performance.now();
        while(performance.now() - start < 10) {}
        return mockData[row - 1][col - 1];
      }
    };
  }
};

const absenceId = 'ID999';
const period = '1'; // subColumnIndex = 10 + 1 - 1 = 10; data index = 9

function original() {
  var data = mockSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(absenceId)) {
      var subColumnIndex = 10 + parseInt(period) - 1;
      var assignedSub = String(mockSheet.getRange(i + 1, subColumnIndex).getValue() || "").trim();
      return assignedSub;
    }
  }
}

function optimized() {
  var data = mockSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(absenceId)) {
      var subColumnIndex = 10 + parseInt(period) - 1;
      var assignedSub = String(data[i][subColumnIndex - 1] || "").trim();
      return assignedSub;
    }
  }
}

console.log("Running original...");
let t0 = performance.now();
let res1 = original();
let t1 = performance.now();
console.log(`Original Time: ${(t1 - t0).toFixed(2)} ms, Result: ${res1}`);

console.log("Running optimized...");
let t2 = performance.now();
let res2 = optimized();
let t3 = performance.now();
console.log(`Optimized Time: ${(t3 - t2).toFixed(2)} ms, Result: ${res2}`);
