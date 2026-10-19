// Mock sheet data
let sheetData = [
  ["Email Mode", "Live"],
  ["App URL", "https://example.com"]
];

function updateSettingsMock(newSettings) {
  var settingsMap = {};

  for (var i = 1; i < sheetData.length; i++) {
    if (!sheetData[i]) continue;
    settingsMap[String(sheetData[i][0]).trim()] = i;
  }

  var rowsToAppend = [];
  var dataChanged = false;

  for (var key in newSettings) {
    if (settingsMap[key] !== undefined) {
      if (sheetData[settingsMap[key]][1] !== newSettings[key]) {
        sheetData[settingsMap[key]][1] = newSettings[key];
        dataChanged = true;
      }
    } else {
      rowsToAppend.push([key, newSettings[key]]);
    }
  }

  console.log("dataChanged:", dataChanged);
  console.log("rowsToAppend:", rowsToAppend);
  console.log("updated sheetData length:", sheetData.length);

  // mock append
  if (rowsToAppend.length > 0) {
     for(var j=0; j<rowsToAppend.length; j++) {
         sheetData.push(rowsToAppend[j]);
     }
  }
  return { success: true };
}

// 1. Send special pay rates
let newSettings = {
  "Green Day Pay Rate": "10",
  "Blue/Gold Day Pay Rate": "20",
  "Special Pay Rates": JSON.stringify([{date: "2026-10-09", rate: "25"}])
};

console.log("Calling updateSettingsMock...");
updateSettingsMock(newSettings);
console.log("SheetData after:", sheetData);
