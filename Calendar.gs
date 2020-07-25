var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();

var properties = PropertiesService.getScriptProperties();

var calendarId = "ki8ogsnb2ibsfh1377di9gf0kc@group.calendar.google.com";



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

  
  var attr = getAttr(x);
  
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
  var user = getKey(y);
  var name = "出勤可能時間（"+user+"）";
 
  var ret = Browser.msgBox(user+"さんにGoogleカレンダーを送りますか?", Browser.Buttons.OK_CANCEL);
  if (ret=='cancel') return;

  var idCell = getAttrValueCell(y, "カレンダーID");
  if (idCell.isBlank()){ //空なら作成
    
    var calendar = CalendarApp.createCalendar(name);
    var id = calendar.getId();
    
    idCell.setValue(id);  
    
    var urlCell = getAttrValueCell(y, "カレンダーURL");
    var url = "https://calendar.google.com/calendar/embed?src="+id;
    urlCell.setValue(url);

  } 
  
  var id = idCell.getValue();
  var calendar = CalendarApp.getCalendarById(id);
  var email = getAttrValueCell(y, "gmail").getValue();
  
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

  var name = "出勤可能時間（"+getKey(y)+"）";
  var ret = Browser.msgBox("カレンダー「"+name+"」を削除しますか? この操作は元に戻せません。", Browser.Buttons.OK_CANCEL);
  if (ret=='cancel') {
    return;
  }

  
  var idCell = getAttrValueCell(y, "カレンダーID");
  var id = idCell.getValue();
  
  var calendar = CalendarApp.getCalendarById(id);
  calendar.deleteCalendar();
  
  idCell.clearContent();

  var urlCell = getAttrValueCell(y, "カレンダーURL");
  urlCell.clearContent();
  
  
}


/** カレンダーが編集されたとき **/
function onCalendarEdit() {
  //このトリガーの場合はカレンダーをIDで開く必要がある。
  var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1qzvi4JiZBuEEgLf8GHVsS3KLRIyH1e9ywK-HoXWYw7g/edit#gid=951613234");
  var events = Calendar.Events.list(calendarId, getSyncToken(calendarId));
 
  var items = events.items;
  
  for (var i = 0; i < items.length; i++) {
    // イベントID
    var eventId = items[i].iCalUID;
    // イベントのタイトル
    var title = items[i].summary;
    // 場所
    var location = items[i].location;
    // 開始日
    var startDate = formatDate(items[i].start);
    // 開始時間
    var startTime = formatTime(items[i].start);
    // 終了時間
    var endTime = formatTime(items[i].end);
    // ステータス
    var status = items[i].status;
    
    // TODO ここにスプレッドシートの内容を更新する処理を書こう
    Logger.log(title);
    Logger.log(startDate);
    Logger.log(startTime);
    addSchedule(items[i]);
                     
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
  

function addSchedule(event){
  
  var sheet = ss.getSheetByName("スケジュール");
  var y = sheet.getLastRow();
  var x = 5;
  Logger.log(y);
  
  var range = sheet.getRange(y+1,x, 1, 5);//次の行
  range.setValues([[
    "",
    formatDate(event.start),
    formatTime(event.start),
    formatTime(event.end),
    ""
  ]])
  
}


/** 
* Debug用。ダイアログを表示。
**/
function msg(value){
  Browser.msgBox(value);
}

