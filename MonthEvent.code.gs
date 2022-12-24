function CopyToAllEvent() {
  
    var ss = SpreadsheetApp.openById("1xuLqZKGMftfKYwaFFaUEuBq44nv2GRUDeWJznQ30EHA")//open allevent
   // var ss = SpreadsheetApp.openById("1s5-W9tGbxb-Jz_3-UnVawtuTofnlogD2CwyHh1eGti4")//open copy All events AYLUS_20221223
    var sheet = ss.getSheetByName("Event Participations");
    var lastRow=sheet.getLastRow() ;
    Logger.log(sheet.getLastRow() + " Is the last Row in All Event.");

    var curSheet = SpreadsheetApp.getActiveSheet();
    Logger.log(curSheet.getLastRow() + " Is the last Row in current sheet.");

    var range = curSheet.getRange(17,1,curSheet.getLastRow()-17+1,3);
    var values = range.getValues();

    values.forEach(function(row, rowId) {
  //row.forEach(function(col, colId) {
      var cell_fname =sheet.getRange(lastRow+rowId+2,1);
      cell_fname.setValue(values[rowId][0]);
      //Logger.log(rowId);
      var cell_lname =sheet.getRange(lastRow+rowId+2,2);
      cell_lname.setValue(values[rowId][1]);

      var event_name=sheet.getRange(lastRow+rowId+2,3);
      event_name.setValue(SpreadsheetApp.getActiveSheet().getName());

      var cell_eventtime =sheet.getRange(lastRow+rowId+2,4);
      cell_eventtime.setValue(SpreadsheetApp.getActiveSheet().getRange(4,2).getValue());            

      var cell_hours =sheet.getRange(lastRow+rowId+2,5);
      cell_hours.setValue(values[rowId][2]);


      Logger.log(values[rowId][0] + values[rowId][1] + values[rowId][2]);
  //});
    });

    //range.setValues(values);

}
