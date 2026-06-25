var payPeriodsData = [
  ["Number", "Start", "End"],
  [1, "2026-06-11", "2026-06-25"]
];

var payPeriods = [];
for (var p = 1; p < payPeriodsData.length; p++) {
  var periodNum = String(payPeriodsData[p][0]).trim();
  var startDateRaw = payPeriodsData[p][1];
  var endDateRaw = payPeriodsData[p][2];

  if (periodNum && startDateRaw && endDateRaw) {
    var startFormatted = String(startDateRaw);
    var endFormatted = String(endDateRaw);
    payPeriods.push({
      periodNumber: periodNum,
      startDate: startFormatted,
      endDate: endFormatted
    });
  }
}
console.log(payPeriods);
