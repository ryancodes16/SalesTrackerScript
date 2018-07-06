//use this function below to give me feedback and suggestions on this project (it actually works too :) )
function giveFeedBack()
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
   var name = "";
   var data = sheet.getDataRange().getValues();
   var NumStore = [];
   var NameStore = [];
   var DateStore = [];
   var QuantityStore = [];
   var check = false;
   var test;
   var type = [];
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
             if(str.substring(0,5) === type[k])
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
  
  //PartNo = ["ES002", "ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021"];
  for(var k = 0; k < PartNo.length; k++)
   {
     name = PartNo[k];
     storeProducts(name);
     
   }
}


function storeProducts(name) {
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
    sheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity']);
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

function DeleteSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //Once again I leveraged the fact that I already made an array of all the part no above, reused this code for what sheets to delete when clearing screen
  var sheets = ["ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021", "Engin", "Hands","415-0", "17032", "ASVTX", "A10X-", "9SIAA","GW541", "TPE-1", "Extra"]; 
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


function createExtra()
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
function addExtra(date, name, num, quantity){   
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("Extra");
  sheet.appendRow([date, name,num, quantity]);
}
  

function sortES002(){
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
    var yourNewSheet = activeSpreadsheet.getSheetByName("ES002-Sorted");
    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }

    yourNewSheet = activeSpreadsheet.insertSheet();
    yourNewSheet.setName("ES002-Sorted");
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.appendRow(['Order Rec\'D Date', 'Cus Name', 'Part No', 'Quantity']);
  for(var i = 1; i < data.length; i++) {
    test = data[i][2].valueOf().substring(0,5);
    if(test === "ES002")
    {
       DateStore[count] = data[i][0];
       NameStore[count] = data[i][1].valueOf();
       NumStore[count] = data[i][2].valueOf();
       QuantityStore[count] = data[i][3].valueOf();
       count++;
    }
   }

  var Sorted = [];
  var text;
  var integer;
  for(var i = 0; i < count; i++)
  {
	text = NumStore[count];
    	var j = count - 1;
    while (j >= 0 && NumStore[j] > text) {
      NumStore[j + 1] = NumStore[j];
      j--;
    }
    NumStore[j + 1] = text;	
  }
  var writeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ES002-Sorted");
  for(i = 0; i < count; i++)
  { 
      writeSheet.appendRow([DateStore[i], NameStore[i],  NumStore[i], QuantityStore[i]]);
  }
}
