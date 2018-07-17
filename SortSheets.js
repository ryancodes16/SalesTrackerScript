function moveSheets_ascending(){
  moveSheets( true );
}
function moveSheets_descending(){
  moveSheets( false );
}

var moveSheets = function ( bySort ){
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = getSheets( bySort );
  for ( var i in sheets ) {
    book.setActiveSheet( book.getSheetByName( sheets[i].getSheetName() ) );
    book.moveActiveSheet( sheets.length );
    Logger.log( sheets[i].getSheetName() + i);
  }
}

var getSheets = function ( bySort ) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.sort(sortFunction);
  if ( !bySort ) { sheets.reverse(); }
  return sheets;
}

var sortFunction = function ( a, b ) {  
  return (a.getSheetName().toUpperCase() < b.getSheetName().toUpperCase()) ? -1 : 1;
}
