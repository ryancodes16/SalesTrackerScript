function addSalesTab(name){
  //var name = "ES004";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(name);
  var data = s.getDataRange().getValues();
  var count = 0;
  var range;
  for(var i = 1; i < data.length; i++){
    Logger.log(data[i][5]);
    if(data[i][5] > 0){
    count += data[i][5].valueOf();
    }
  }
  Logger.log("$" + count);
  Logger.log(data.length);
  range = s.getRange("F" + (data.length));
  range.setValue("$" + count); 
}
