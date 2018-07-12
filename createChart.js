//simple function

function newChart() {
  // Generate a chart representing the data in the range of A1:B15.
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Test");

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.BAR)
     .addRange(sheet.getRange('A1:B15'))
     .setPosition(5, 5, 0, 0)
     .build();

  sheet.insertChart(chart);
}
