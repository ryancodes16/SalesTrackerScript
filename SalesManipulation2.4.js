//use this function below to give me feedback and suggestions on this project (it actually works too :) )
function giveFeedBack() //works perfectly
{
  var emailAddress = "regier1959@students.d211.org";
  var prompt = "Any suggestions or constructive criticism?";
  var message = Browser.inputBox(prompt);       
  var subject = "Epiq Project Feedback";
  MailApp.sendEmail(emailAddress, subject, message);
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

// ****RUN ME ****
function STARTPROGRAM() //use this to start the program (autogenerates the array of Part Prefixes which can then be ran through the sorting method) *******ONLY RUN THIS TO START THE PROGRAM*******
{
  var PartNo = [];
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("From_MRP");
  var data = sheet.getDataRange().getValues(); 
  var check = false;
  var index = 0;
  var str;
  /*FIXME We need to check the p/n's and automatically populate the array below rather than hard-code them
  First check the first two characters, if NOT "ES" then this part number will be pushed into "Extras" tab
  If "ES" number, then check the next 3 characters and if they are unique then add to the array list.
  Ideally this array list will be sorted after exiting the function.                                        ***********************UPDATE: ISSUE FIXED 07/6/18****************************
  */
  createExtra();
  for(var i=1; i< data.length; i++)
  {
    check = true;
    str = data[i][2].valueOf();
    if(str.substring(0,2) === "ES")
    {
      for(var k = 0; k < PartNo.length; k++)
      {
        if(str.substring(0,5) === PartNo[k])
        {
          check = false;
        }             
      }
      if(check === true)
      {
        PartNo[index] = str.substring(0,5);
        index++;
      }
    }
    else
    {
      addExtra(data[i][0], data[i][1].valueOf(), data[i][2].valueOf(), data[i][3].valueOf());
    }
  }
  
  //PartNo = ["ES002", "ES011", "ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021"]; //old manual array, using automated one now
  for(var k = 0; k < PartNo.length; k++)
  {
    name = PartNo[k];
    storeProducts(name);
    subtotals(name);
  }
  salesTab1();
}


function storeProducts(name) { //works perfectly, creates indivdual sheets from array of part numbers, sorts parts in indivudal sheets by last three digits of product ID, creates sub-totals for parts with same 
  //last three digits in Product ID
  //name = "ES002";
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("From_MRP");
  var data = sheet.getDataRange().getValues();
  var NumStore = [];
  var NameStore = [];
  var DateStore = [];
  var ES002 = [];
  var QuantityStore = [];
  var count = 0;
  var test;
  var date;
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var yourNewSheet = activeSpreadsheet.getSheetByName(name);
  if (yourNewSheet != null) {
    activeSpreadsheet.deleteSheet(yourNewSheet);
  }
  
  yourNewSheet = activeSpreadsheet.insertSheet();
  yourNewSheet.setName(name);
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity', 'Sub-Totals', 'Sales', 'Average Price']);
  format(name);
  
  for(var i = 1; i < data.length; i++) {
    test = data[i][2].valueOf().substring(0,5);
    if(test === name)
    {
      ES002[count] = data[i];
      DateStore[count] = data[i][0];
      NameStore[count] = data[i][1].valueOf();
      NumStore[count] = data[i][2].valueOf();
      QuantityStore[count] = data[i][3].valueOf();
      count++;
    }
  }
  
  var tempNum, tempDate, tempName, tempQuantity;
  for(var i = 0; i < (count-1); i++)
  {
    for(var j = 0; j < (count-i-1); j++)
    {
      if(NumStore[j].substring(6,10) > NumStore[j+1].substring(6,10))
      {
        tempNum = NumStore[j];
        NumStore[j] = NumStore[j+1];
        NumStore[j+1] = tempNum;
        
        tempDate = DateStore[j];
        DateStore[j] = DateStore[j+1];
        DateStore[j+1] = tempDate;
        
        tempName = NameStore[j];
        NameStore[j] = NameStore[j+1];
        NameStore[j+1] = tempName;
        
        tempQuantity = QuantityStore[j];
        QuantityStore[j] = QuantityStore[j+1];
        QuantityStore[j+1] = tempQuantity;
      }
    }
  }
  
  
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  for(i = 0; i < count; i++)
  { 
    writeSheet.appendRow([DateStore[i], NameStore[i],  NumStore[i], QuantityStore[i]]);
  }
}


function DeleteSheets(){ //Delete all existing sheets besides the MRP ******UPDATE TO BE AUTOMATED******
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //Once again I leveraged the fact that I already made an array of all the part no above, reused this code for what sheets to delete when clearing screen
  var sheets = ["ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021", "ES011", "Extra", "ES002"]; 
  var numberOfSheets = ss.getSheets().length;
  for(var s = numberOfSheets-1; s>=0 ; s--){
    SpreadsheetApp.setActiveSheet(ss.getSheets()[s]);
    var shName = SpreadsheetApp.getActiveSheet().getName();
    
    if(sheets.indexOf(shName)>-1){
      var delSheet = ss.deleteActiveSheet();
      Utilities.sleep(500);
    }
  }
  SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);// send me back to first sheet (original list of sales
}


function createExtra() //generates 'Extra' tab
{
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("From_MRP");
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var yourNewSheet = activeSpreadsheet.getSheetByName("Extra");
  if (yourNewSheet != null) {
    activeSpreadsheet.deleteSheet(yourNewSheet);
  }
  yourNewSheet = activeSpreadsheet.insertSheet();
  yourNewSheet.setName("Extra");
  yourNewSheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity', 'Average Price']);
  format("Extra");
}


function addExtra(date, name, num, quantity){   //adds items who aren't 'ES ' parts to the 'Extra' tab
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("Extra");
  sheet.appendRow([date, name,num, quantity]);
}




function subtotals(name){ //will display sub-totals for each individual product's sheets, add quantity from each indivudal three digit ending of product ID
  Logger.clear();
  Logger.log(name);
  //name = "ES002";
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(name);
  var data = sheet.getDataRange().getValues();
  var temp, temp2;
  var count;
  var range;
  //Logger.log(data);
  //range = sheet.getRange("E" + (data.length + 2)); 
  //range.setValue(count);
  for(var i = 1; i < data.length + 1; i++)
  {
    Logger.log(i + "i");
    if(i === 1)
    {
      count = 0;
      count = data[i][3].valueOf();
      //Logger.log("Count for " + data[i][2] + ": " + count);
    }
    else if(i === data.length)
    {
      Logger.log(count + " final count");
      range = sheet.getRange("E" + (i)); 
      range.setValue(count);
    }
    else
    {
      Logger.log(data[i] + " i = " + i);
      //Logger.log("#1: " + data[i][2] + " #2: " + data[i-1][2]);
      if(data[i][2] === data[i-1][2])
      {
        //Logger.log("Yes");
        count += data[i][3];
        //Logger.log("Count for " + data[i][2] + ": " + count);
      }
      else
      {
        //Logger.log("No");
        //Logger.log("Count for " + data[i-1][2] + ": " + count);
        Logger.log(count + " count");
        range = sheet.getRange("E" + (i)); 
        range.setValue(count);
        count = data[i][3];
      }
    }
    
  }
  
  count = 0;
  for(var i = 1; i < data.length; i++){
    count += data[i][3];
    Logger.log(count);
  }
  sheet.appendRow(["7/09/2018", "Epiq Solutions", "Ryan Regier", "Total: " + count, "Total: " + count])
}

function format(name){ //formats header at top of each sheet (bolds it and freezes it)
  //name = "ES002";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(name);
  if(name === "Extra")
  {
    s.getRange(1, 1).setFontWeight("bold");
    s.getRange(1, 2).setFontWeight("bold");
    s.getRange(1, 3).setFontWeight("bold");
    s.getRange(1, 4).setFontWeight("bold");
    s.getRange(1, 5).setFontWeight("bold");
  }
  else
  {
    s.getRange(1, 1).setFontWeight("bold");
    s.getRange(1, 2).setFontWeight("bold");
    s.getRange(1, 3).setFontWeight("bold");
    s.getRange(1, 4).setFontWeight("bold");
    s.getRange(1, 5).setFontWeight("bold");
    s.getRange(1, 6).setFontWeight("bold");
    s.getRange(1, 7).setFontWeight("bold");
  }
  s.setFrozenRows(1);
}



function salesTab1(){ //updated salesTab to work for multiple sales for different parts inputted into quickbooks
  Logger.clear();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("From_QBs");
  var QBdata = s.getDataRange().getValues();
  var partData;
  var temp, range;
  //Logger.log(QBdata);
  for(var z = 1; z < QBdata.length; z++){
  var name = QBdata[z][0].substring(0,9).valueOf();
  Logger.log(name);
  Logger.log(name.substring(0,5));
  s = ss.getSheetByName(name.substring(0,5));
  partData = s.getDataRange().getValues();
  //Logger.log("$" + QBdata[z][2].valueOf() / QBdata[z][1].valueOf());
  for(var i = 1; i < partData.length; i++){
    Logger.log("I " + i);
    if(i === 1)
      temp++;
    if(partData[i][2].toString().trim() == name.toString().trim()) {   // force convert to string
      temp = i + 1; 
    }
  }
  Logger.log(temp);
  range = s.getRange("F" + temp); 
  range.setValue("$" + QBdata[z][2].valueOf());
  Logger.log("SAlES: " + QBdata[z][2].valueOf());
  
  var range2 = s.getRange("E" + temp);
  var numSold = range2.getValue();
  Logger.log(numSold + " sold");
  range = s.getRange("G" + temp); 
  range.setValue("$" + ((QBdata[z][2].valueOf()-QBdata[z][3]) / numSold));
    }
}

  





function salesTab(){ //updated salesTab to work for one part/sales for ES002 parts inputted into quickbooks
  Logger.clear();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("From_QBs");
  var QBdata = s.getDataRange().getValues();
  var partData;
  var temp, range;
  //Logger.log(QBdata);
  for(var z = 1; z < QBdata.length; z++){
  var name = QBdata[z][0].substring(0,9).valueOf();
  if(name.substring(0,5) === "ES002"){
  Logger.log(name);
  Logger.log(name.substring(0,5));
  s = ss.getSheetByName(name.substring(0,5));
  partData = s.getDataRange().getValues();
  //Logger.log("$" + QBdata[z][2].valueOf() / QBdata[z][1].valueOf());
  for(var i = 1; i < partData.length; i++){
    Logger.log("I " + i);
    if(i === 1)
      temp++;
    if(partData[i][2].toString().trim() == name.toString().trim()) {   // force convert to string
      temp = i + 1; 
    }
  }
  Logger.log(temp);
  range = s.getRange("F" + temp); 
  range.setValue("$" + QBdata[z][2].valueOf());
  Logger.log("SAlES: " + QBdata[z][2].valueOf());
  
  var range2 = s.getRange("E" + temp);
  var numSold = range2.getValue();
  Logger.log(numSold + " sold");
  range = s.getRange("G" + temp); 
  range.setValue("$" + ((QBdata[z][2].valueOf() / numSold)-QBdata[z][3]));
    }
  }
}




























function viewUser(){ //experimenting with charts in google scripts
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName("From_MRP");
  var me = Session.getEffectiveUser();
  
  Logger.log(me);
}

function emailResults(name){
  //works perfectly
  name = "ES002";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(name);
  var data = s.getDataRange().getValues();
  var emailAddress = "regier1959@students.d211.org";
  var message = data;       
  var subject = "Data from: " + name;
  MailApp.sendEmail(emailAddress, subject, message);
}

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
