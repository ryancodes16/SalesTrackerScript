//use this function below to give me feedback and suggestions on this project (it actually works too :) )
function giveFeedBack() //works perfectly
{
  var emailAddress = "regier1959@students.d211.org";
  var prompt = "Any suggestions or constructive criticism?";
  var message = Browser.inputBox(prompt);       
  var subject = "Epiq Project Feedback";
  MailApp.sendEmail(emailAddress, subject, message);
}

function runMe()
{
  var PartNo = [];
  createExtra();
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("From_MRP");
  var data = sheet.getDataRange().getValues(); 
  var check = false;
  var index = 0;
  var str;
  /*FIXME We need to check the p/n's and automatically populate the array below rather than hard-code them
  First check the first two characters, if NOT "ES" then this part number will be pushed into "Extras" tab
  If "ES" number, then check the next 3 characters and if they are unique then add to the array list.
  Ideally this array list will be sorted after exiting the function.
  */
  for(var i=0; i< data.length; i++)
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
    
  }
}


function storeProducts(name) { //works perfectly
  //name = "ES002";
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("From_MRP");
  var data = sheet.getDataRange().getValues();
  var NumStore = [];
  var NameStore = [];
  var DateStore = [];
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
  sheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity', 'Sub-Totals']);
  for(var i = 1; i < data.length; i++) {
    test = data[i][2].valueOf().substring(0,5);
    if(test === name)
    {
      DateStore[count] = data[i][0];
      NameStore[count] = data[i][1].valueOf();
      NumStore[count] = data[i][2].valueOf();
      QuantityStore[count] = data[i][3].valueOf();
      count++;
    }
  }
  
  
  
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  for(i = 0; i < count; i++)
  { 
    writeSheet.appendRow([DateStore[i], NameStore[i],  NumStore[i], QuantityStore[i]]);
  }
  
}

function DeleteSheets(){ //works perfectly
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //Once again I leveraged the fact that I already made an array of all the part no above, reused this code for what sheets to delete when clearing screen
  var sheets = ["ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021", "ES011", "Extra"]; 
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


function createExtra() //works perfectly
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
  yourNewSheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity']);
}
function addExtra(date, name, num, quantity){   //works perfectly
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("Extra");
  sheet.appendRow([date, name,num, quantity]);
}

function sortStuff(name){
  name = "ES002";  
  var PartType = name;
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(PartType);
  var data = sheet.getDataRange().getValues(); 
  var nameTemp, partNoTemp, dateTemp, quantityTemp;
  
  var n = data.length;
  for(var i = 0; i < n; i++)
  {
    for(var j = 1; j < (n-i); j++)
    {
      if(data[j-1][2].valueOf().substring(5,8) > data[j][2].valueOf().substring(5,8))
      {
        partNoTemp = data[j-1][2].valueOf();
        data[j-1][2] = data[j][2].valueOf();
        data[j][2] = partNoTemp;
        
      }
      
    }
    
    
  }
    
  
}






function storeProducts2(name) { //works perfectly
  name = "ES002";
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
  sheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity', 'Sub-Totals']);
  
  
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

function subtotals(){
  
  
}
