
function getColumnIndex(columnName){
  var headers = SpreadsheetApp.getActiveSpreadsheet().getDataRange().getValues().shift();
  var colindex = headers.indexOf(columnName);
  return colindex+1;
}

function dateFormat(whichDate, whatTime){

  //'september 21, 2022 14:00:00'
  var dateString = Utilities.formatDate(new Date(whichDate),'America/New_York',"MMMM dd, yyyy");
  var timeString = Utilities.formatDate(new Date(whatTime),'America/New_York',"HH:mm:ss");
  var retVal = dateString +" "+ timeString;
  SpreadsheetApp.getUi().alert(retVal);
  return new Date(retVal);
}
function CreatSheet() {

    var s_all_events = SpreadsheetApp.getActiveSheet();
    var cell_active = s_all_events.getActiveCell();
    var rowi = cell_active.getRowIndex();
    var coli = getColumnIndex("DONOTCHANGE_SheetId");

    var cell_sheetid = s_all_events.getRange(cell_active.getRowIndex(), getColumnIndex("DONOTCHANGE_SheetId"));    
    var cell_tabname = s_all_events.getRange(cell_active.getRowIndex(), getColumnIndex("DONOTCHANGE_TabName"));  

     var sheetid = cell_sheetid.getValue();
     var tabname = cell_tabname.getValue();
    var s_month_event=SpreadsheetApp.openById(sheetid);    
    var target_tab = s_month_event.getSheetByName(tabname);
    if(target_tab==null)
    {
    var target_tab = s_month_event.insertSheet(tabname,0, {template: s_month_event.getSheetByName("9/")});
      
    }

    var cell_event_name_from = s_all_events.getRange(cell_active.getRowIndex(), getColumnIndex("Event Name"));  
    var cell_event_name_to = target_tab.getRange(2, 2);
    var cell_event_name_from_val=cell_event_name_from.getValue();
    cell_event_name_to.setValue(cell_event_name_from.getValue());
    //Browser.msgBox(eventNameCell.getFormula());
    //Browser.msgBox(eventNameCell.getFormula().replace("18","21"));
    
   // eventNameCell.setValue(eventNameCell.getFormula().replaceAll("18","21"));
  //var selection = s.getSelection().getActiveRange().getValues();
  //Browser.msgBox(selection);
  //selection.forEach(function(entey){

  //});
}

function TransferToCalendar() {
    var s = SpreadsheetApp.getActiveSheet();
    var selection = s.getSelection().getActiveRange().getValues();
    //Browser.msgBox(selection);
}

function UpdateCalendar() {

 SpreadsheetApp.getUi().alert('Hello, world');
 try {
  var s = SpreadsheetApp.getActiveSheet();

  var cell = s.getActiveCell();
  //var nextCell = cell.offset(0,1);
  //nextCell.setValue(new Date()).setNumberFormat("yyyy-MM-dd hh:mm");
  var rowIndex = cell.getRowIndex();

  if(cell.getRowIndex() > 1){

    //var nextCell = datecell.offset(0,1);
    //nextCell.setValue(rowIndex);

    // var rows = s.getRange(cell.getRowIndex()).getValues();
    SpreadsheetApp.getUi().alert(rowIndex);
    let aylusCalendar = CalendarApp.getCalendarById("ayluscheshire@gmail.com");
    //SpreadsheetApp.getUi().alert('1');

    var calendrIdCell = s.getRange(cell.getRowIndex(), getColumnIndex("CalendarEventId"));
    //SpreadsheetApp.getUi().alert('2');

      var eventNameCell = s.getRange(cell.getRowIndex(), getColumnIndex("Event Name"));
      var startTimeCell = s.getRange(cell.getRowIndex(), getColumnIndex("Start Time"));
      var endTimeCell = s.getRange(cell.getRowIndex(), getColumnIndex("End Time"));
      var eventDateCell = s.getRange(cell.getRowIndex(), getColumnIndex("Event Date"));
      
      // delete existing one 
      if(calendrIdCell.getValue() !="")
      {
        let existingEvent = aylusCalendar.getEventById(calendrIdCell.getValue());
        existingEvent.deleteEvent();
      }

      // insert new event
        //var event=aylusCalendar.createEvent("test111", new Date('september 21, 2022 14:00:00'), new Date('september 21, 2022 15:00:00'));
        var event=aylusCalendar.createEvent(eventNameCell.getValue(), dateFormat(eventDateCell.getValue(), startTimeCell.getValue()), dateFormat(eventDateCell.getValue(), endTimeCell.getValue()));
        calendrIdCell.setValue(event.getId());
          Logger.log(event.getId());
  }
 }
catch (err) {
    // return the error object so we know where we are
    SpreadsheetApp.getUi().alert(err);
    //var stack = err.stack.split('\n');
}

  //Logger.log(e.range.getA1Notation());
}

