function topCust(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("Summary");
  var a = sheet.getDataRange().getValues();
  sheet.appendRow([' ']);
  sheet.appendRow(["Top Paying Customers", "Name", "Quantity", "Sales"]);
  var tempQ, tempN, tempM
  var swapped;
    do {
        swapped = false;
        for (var i=1; i < a.length-1; i++) {
            if (a[i][1] > a[i+1][1]) {
                tempQ = a[i][1];
                a[i][1] = a[i+1][1];
                a[i+1][1] = tempQ;
                 
                tempN = a[i][0];
                a[i][0] = a[i+1][0];
                a[i+1][0] = tempN;
              
              
                tempM = a[i][2];
                a[i][2] = a[i+1][2];
                a[i+1][2] = tempM;
                Logger.log(tempM);
                Logger.log(tempN);
                Logger.log(tempQ);
                swapped = true;
            }
        }
    } while (swapped);
  for(var z = a.length - 1; z > (a.length - 1) - 5; z--){
    Logger.log(a[z][2]);
    Logger.log(a[z][1]);
    Logger.log(a[z][0]);
    sheet.appendRow([a[z][0], a[z][1], a[z][2]]);
  }
}
