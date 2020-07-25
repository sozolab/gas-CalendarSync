var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

var properties = PropertiesService.getScriptProperties();



/**
 * スプレッドシート表示の際に呼出し
 */
function onOpen() {
  
  var subMenus = [];
  subMenus.push({
    name: "カレンダー出力",
    functionName: "export2Calendar" 
  });
  
  ss.addMenu("カレンダー", subMenus);
  
}


function onEditTrigger(e){
  var y = e.range.getRow();
  var x = e.range.getColumn();
  var cell = e.range;

  
  var attr = getAttr(sheet, x);
  
  if (attr=="カレンダー送信"){
    if (cell.getValue()==true){
      setStaffCalendar(y,x, cell);
    }
  } else if (attr=="カレンダー削除"){
    if (cell.getValue()==true){
      deleteStaffCalendar(y,x, cell);
    }
  }
}

function setStaffCalendar(y,x, cell){
  cell.setValue(false);
  var user = getKey(sheet, y);
  var name = "出勤可能時間（"+user+"）";
 
  var ret = Browser.msgBox(user+"さんにGoogleカレンダーを送りますか?", Browser.Buttons.OK_CANCEL);
  if (ret=='cancel') return;

  var idCell = getAttrValueCell(sheet, y, "カレンダーID");
  if (idCell.isBlank()){ //空なら作成
    
    var calendar = CalendarApp.createCalendar(name);
    var id = calendar.getId();
    
    setCalendarTrigger(id);
    
    idCell.setValue(id);  
    
    var urlCell = getAttrValueCell(sheet, y, "カレンダーURL");
    var url = "https://calendar.google.com/calendar/embed?src="+id;
    urlCell.setValue(url);

  } 
  
  var id = idCell.getValue();
  var calendar = CalendarApp.getCalendarById(id);
  var email = getAttrValueCell(sheet, y, "gmail").getValue();
  
  shareStaffCalendar(calendar, email);
  
}

//calendar をemailに編集権限ありで共有
function shareStaffCalendar(calendar, email){
  calendar.setSelected(true);
  
  Calendar.Acl.insert({
    "role": "writer",
    "scope": {
      "type": "user",
      "value": email
    }
  }, calendar.getId());
}
    

function deleteStaffCalendar(y,x, cell){
  cell.setValue(false);

  var name = "出勤可能時間（"+getKey(sheet, y)+"）";
  var ret = Browser.msgBox("カレンダー「"+name+"」を削除しますか? この操作は元に戻せません。", Browser.Buttons.OK_CANCEL);
  if (ret=='cancel') {
    return;
  }

  
  var idCell = getAttrValueCell(sheet, y, "カレンダーID");
  var id = idCell.getValue();
  
  var calendar = CalendarApp.getCalendarById(id);
  if (calendar != null) calendar.deleteCalendar();
  
  deleteCalendarTrigger(id);
  
  idCell.clearContent();

  var urlCell = getAttrValueCell(sheet, y, "カレンダーURL");
  urlCell.clearContent();
  
  
}


/** カレンダーが編集されたとき **/
function onCalendarEdit(e) {
  //このトリガーの場合はカレンダーをIDで開く必要がある。
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1qzvi4JiZBuEEgLf8GHVsS3KLRIyH1e9ywK-HoXWYw7g/edit#gid=951613234");
  var calendarId = e.calendarId;
  Logger.log(calendarId);
  var events = Calendar.Events.list(calendarId, getSyncToken(calendarId));
 
  var items = events.items;
  
  for (var i = 0; i < items.length; i++) {
    // 場所
    var location = items[i].location;
    // ステータス
    var status = items[i].status;
    
    updateSchedule(items[i]);
                     
  }
  
  setSyncToken(events, calendarId);
}


/** Initial Sync Token**/
function initialSync(calendarId) {
  var items = Calendar.Events.list(calendarId);
  var nextSyncToken = items.nextSyncToken;
  properties.setProperty(getSyncKey(calendarId), nextSyncToken);
}


function getSyncKey(calendarId){
  return "syncToken:"+calendarId;
}

/** get Sync Token **/
function getSyncToken(calendarId){
  var nextSyncToken = properties.getProperty(getSyncKey(calendarId));
  var optionalArgs = {
    syncToken: nextSyncToken
  };
  return optionalArgs;
}

/** set Sync Token for next **/
function setSyncToken(events, calendarId){

  var nextSyncToken = events["nextSyncToken"];
  properties.setProperty(getSyncKey(calendarId), nextSyncToken);
}


function formatDate(eventDate){
  var date = Utilities.formatDate(new Date(eventDate.dateTime), "Asia/Tokyo", "yyyy/MM/dd");
  return date;
}

function formatTime(eventDate){
  var time = Utilities.formatDate(new Date(eventDate.dateTime), "Asia/Tokyo", "HH:mm");
  return time;
}


var scheduleSheet = ss.getSheetByName("スケジュール");
var eventX = 5;
var eventColSize=7;


function updateSchedule(event){
  var id = event.iCalUID;
  Logger.log("id:"+id);
  var x = getAttrX(scheduleSheet, "イベントID");
  Logger.log("x:"+x);

  var y = getY(scheduleSheet, "イベントID", id);
  Logger.log("y:"+y);
  
  if (y<0){ //if not found
    y = scheduleSheet.getLastRow()+1; //new row
  }
  
  var range = scheduleSheet.getRange(y,eventX, 1, eventColSize);//次の行
  range.setValues([[
    "",
    event.summary,
    id, 
    formatDate(event.start),
    formatTime(event.start),
    formatTime(event.end),
    ""
  ]]);
  
  
}





/** 
* Debug用。ダイアログを表示。
**/
function msg(value){
  Browser.msgBox(value);
}


//TODO: 権限管理
// カレンダー編集時のトリガーを設定
function setCalendarTrigger(calendarId){

  // これを実行している人のメールアドレス
  var email = Session.getActiveUser().getEmail();
  if( email === "<mail_address>"){
    // 正常処理
  }else{
    // 異常検知処理
//    return ;
  }

  ScriptApp
  .newTrigger("onCalendarEdit")
  .forUserCalendar(calendarId)
  .onEventUpdated()
  .create();

}

function deleteCalendarTrigger(calendarId){
  var triggers = ScriptApp.getProjectTriggers();
  for( var i = 0; i < triggers.length; ++i ){
    var trigger = triggers[i];
    if (trigger.getTriggerSourceId()==calendarId){
      ScriptApp.deleteTrigger(trigger);
    }
  }
}