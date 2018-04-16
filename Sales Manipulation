/use this function below to give me feedback and suggestions on this project (it actually works too :) )
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
   var name = "";
   var sheet = SpreadsheetApp.getActiveSheet();
   var data = sheet.getDataRange().getValues();
   var PartNo = [];
   var check;
  /* Ignore this section, I was trying to have it automatically read in each Part No but it wasn't working so I manually used an array I pre-filled with all the Part No #s
   for(var i = 0; i < data.length; i++)
   {
     check = true;
     for(var z = 0; z < PartNo.length; z++)
     {
        if(data[z][2].valueOf().substring(0,5) === PartNo[i])
        {
           check = false; 
        }       
     }
     if(check === true)
     {
       PartNo[i] = data[z][2].valueOf().substring(0,5); 
       
     }
     
   }
   for(var k = 0; k < PartNo.length; k++)
   {
     name = PartNo[k];
     storeProducts(name);
     
   }
   */
  //Array holding each Part No 
  PartNo = ["ES002", "ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021", "Engin", "Hands","415-0", "17032", "ASVTX", "A10X-", "9SIAA","GW541", "TPE-1"];
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
  var sheets = ["ES002", "ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021", "Engin", "Hands","415-0", "17032", "ASVTX", "A10X-", "9SIAA","GW541", "TPE-1"]; 
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
function findMissing(){
  Logger.clear();
  var check = false;
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  PartNo = ["ES002", "ES003", "ES004" , "ES005", "ES010","ES012","ES013", "ES014", "ES015", "ES016","ES020", "ES021", "Engin", "Hands","415-0", "17032", "ASVTX", "A10X-", "9SIAA","GW541", "TPE-1"];
  for(var i = 0; i < data.length; i++)
  {
    check = false;
    for(var k = 0; k < PartNo.length; k++)
    {
        if(PartNo[k] === data[i][2].valueOf().substring(0,5))
        {
           check = false; 
        }
        else
        {
           check = true; 
        }
    }
    if(check === true)
    {
      Logger.log(data[i][2].valueOf().substring(0,5));  
    }
  }
  
}
