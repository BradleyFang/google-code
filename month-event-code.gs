var DATETIME_HEADER = "UpdatedTime";

function getColumnIndex(columnName){
  var headers = SpreadsheetApp.getActiveSpreadsheet().getDataRange().getValues().shift();
  var colindex = headers.indexOf(columnName);
  return colindex+1;
}

function dateFormat(whichDate, whatTime){

  //'september 21, 2022 14:00:00'
  var dateString = Utilities.formatDate(new Date(whichDate),'America/New_York',"MMMM dd, yyyy");

  var retVal = dateString;

  if(whatTime!="")
  {
    var timeString = Utilities.formatDate(new Date(whatTime),'America/New_York',"HH:mm:ss");
    retVal =retVal + timeString;
  }
  //SpreadsheetApp.getUi().alert(retVal);
  return new Date(retVal);
}

function getDatetimeCol(){
  var headers = SpreadsheetApp.getActiveSpreadsheet().getDataRange().getValues().shift();
  var colindex = headers.indexOf(DATETIME_HEADER);
  return colindex+1;
}
function onEdit(e) {
  var s = SpreadsheetApp.getActiveSheet();
  var cell = s.getActiveCell();
  //var nextCell = cell.offset(0,1);
  //nextCell.setValue(new Date()).setNumberFormat("yyyy-MM-dd hh:mm");
  var rowIndex = cell.getRowIndex();
  if(cell.getRowIndex() > 1){
    var datecell = s.getRange(cell.getRowIndex(), getDatetimeCol());
    datecell.setValue(new Date()).setNumberFormat("yyyy-MM-dd hh:mm");
   // var nextCell = datecell.offset(0,1);
    //nextCell.setValue(rowIndex);
  }
}
