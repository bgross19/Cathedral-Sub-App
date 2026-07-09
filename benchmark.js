const iterations = 10000;

// Mock Spreadsheet with many sheets to simulate Apps Script behaviour
const sheets = [];
for (let i = 0; i < 50; i++) {
  sheets.push({ getName: () => `Sheet${i}` });
}
sheets.push({ getName: () => "PayPeriods" });
for (let i = 51; i < 100; i++) {
  sheets.push({ getName: () => `Sheet${i}` });
}

const ss = {
  getSheets: () => sheets,
  getSheetByName: (name) => sheets.find(s => s.getName() === name) || null
};

console.log("Benchmarking original implementation...");
const startOrig = Date.now();
for (let i = 0; i < iterations; i++) {
  const allSheets = ss.getSheets();
  let payPeriodsSheet = null;
  for (let s = 0; s < allSheets.length; s++) {
    if (allSheets[s].getName().toLowerCase() === "payperiods") {
      payPeriodsSheet = allSheets[s];
      break;
    }
  }
}
const endOrig = Date.now();
const origTime = endOrig - startOrig;
console.log(`Original: ${origTime} ms`);

console.log("Benchmarking optimized implementation...");
const startOpt = Date.now();
for (let i = 0; i < iterations; i++) {
  let payPeriodsSheet = ss.getSheetByName("PayPeriods");
}
const endOpt = Date.now();
const optTime = endOpt - startOpt;
console.log(`Optimized: ${optTime} ms`);
console.log(`Improvement: ~${((origTime - optTime) / origTime * 100).toFixed(2)}% faster`);
