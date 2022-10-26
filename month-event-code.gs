var DATETIME_HEADER = "UpdatedTime";

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
