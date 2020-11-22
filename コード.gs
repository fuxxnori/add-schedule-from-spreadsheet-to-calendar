function myFunction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheets()[0]
  const calendarId = sh.getRange('A2').getValue()
  const calendar = CalendarApp.getOwnedCalendarById(calendarId);
  const events = getInputEvents(sh);
  events.forEach(function(event){
    const title = event[0];
    const startTime = event[1];
    const endTime = event[2];
    const date = event[3];
    if(date){
      calendar.createAllDayEvent(title, date);
    }else{
      calendar.createEvent(title, startTime, endTime);
    }
  })
}

function getInputEvents(sh){
  const lastRow = sh.getRange(sh.getMaxRows(),1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  if(lastRow<=6){
    Browser.msgBox("予定が入力されていません。")
    return;
  }else{
    const events = sh.getRange(`A7:D${lastRow}`).getValues();
    return events;
    /*
    events are expected an array like below
    [
        ['title','start','end','date'],
        ['title','start','end','date'],
        ['title','start','end','date']
    ]
    */
  }
}
