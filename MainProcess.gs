// ==================================================
// 出退勤チェック
// トリガーの設定でこの関数が呼ばれる
// ==================================================
function checkInput() {
  var manager =  new ParupunteManager(EXECUTE_TYPE);
  manager.processGet();
}

// ==================================================
// 返信チェック
// ==================================================
function checkReply(){
  var manager =  new ParupunteManager(EXECUTE_TYPE);
  manager.processSet();
}

// ==================================================
// 入力漏れのチェック
// トリガーの設定でこの関数が呼ばれる
// ==================================================
function checkInputMiss() {
  var manager =  new ParupunteManager(EXECUTE_TYPE);
  manager.processCheckInputMiss();
}

// ==================================================
// ログに記録する
// ==================================================
function _writeLog(params) {
  var spreadSheet = SpreadsheetApp.openByUrl(LOGBOOK_URL);
  var targetSheet = spreadSheet.getSheetByName(NAME_LOG_SHEET);
  
  // 記入行を取得
  var targetNo = parseInt(targetSheet.getRange("K1").getValue())+1;
  
  //No
  targetSheet.getRange("R"+targetNo+"C1").setValue(targetNo-1);
  
  //記入
  targetSheet.getRange("R"+targetNo+"C3").setValue(_changeFormatTime(new Date(),0));
  
  //時間
  targetSheet.getRange("R"+targetNo+"C4").setValue(params.time);
  
  //roomID
  targetSheet.getRange("R"+targetNo+"C5").setValue(params.roomID);
  
  //messageID
  targetSheet.getRange("R"+targetNo+"C6").setValue(params.messageID);
  
  //accountID
  targetSheet.getRange("R"+targetNo+"C7").setValue(params.accountID);
  
  //userName
  targetSheet.getRange("R"+targetNo+"C8").setValue(params.userName);
  
  //body
  targetSheet.getRange("R"+targetNo+"C9").setValue(params.body);
  
  //type
  targetSheet.getRange("R"+targetNo+"C10").setValue(params.type);

  // 記入行を更新
  targetSheet.getRange("K1").setValue(targetNo);
  
}

// ==================================================
// メッセージを取得しログに記録する
// ==================================================
function _processGet(vAccountToken,vRoomID){
  
  var Params = function(){
    this.time;
    this.roomID;
    this.messageID;
    this.accountID;
    this.userName;
    this.body;
    this.type;
  }

  // ChatWorkからメッセージの取得
  var response = _getMessage(vAccountToken,vRoomID);

  for ( var i in response ){
    
    var params = new Params();

    params.accountID = response[i].account.account_id;
    
    // 管理アカウントはスルー
    if(params.accountID == ACCOUNT_ID_DQMANAGER)continue;
    if(params.accountID == ACCOUNT_ID_PARUPUNTA)continue;
    
    params.time = _changeFormatTime(new Date( response[i].send_time * 1000 ),0);
    params.roomID = vRoomID;
    params.messageID = response[i].message_id;
    params.userName = response[i].account.name;
    params.body = response[i].body;
    
    // メッセージ内容からタイプを判定
    params.type = _judgeMethod(response[i].body);
    
    _writeLog(params);
  }
}

// ==================================================
// ログから返信を行う
// ==================================================
function _processSet(vAccountToken,vProcessType){
  
  var Params = function(){
    this.check;
    this.input;
    this.time;
    this.roomID;
    this.messageID;
    this.accountID;
    this.userName;
    this.body;
    this.type;
    this.result;
    this.serif;
  }
  
  var userList = _getUserAll();
 
  var spreadSheet = SpreadsheetApp.openByUrl(LOGBOOK_URL);
  var targetSheet = spreadSheet.getSheetByName(NAME_LOG_SHEET);
  
  // 記録数と返信数の差分件数だけループさせる
  var writeNo = parseInt(targetSheet.getRange("K1").getValue());
  var replyNo = parseInt(targetSheet.getRange("L1").getValue());
  
  var processNum = writeNo - replyNo;
  
  // 
  for ( var i = 0 ; i < processNum ; i++){
    
    var params = new Params();
    
    var targetNo = parseInt(targetSheet.getRange("L1").getValue())+1;
    
    //構造体的なものに詰め込む
    
    //返信
    params.check = targetSheet.getRange("R"+targetNo+"C2").getValue();
    if(params.check != "")continue;
    
    //記入
    params.input = targetSheet.getRange("R"+targetNo+"C3").getValue();
    if(params.input == "")break; 
    
    //時間
    params.time = targetSheet.getRange("R"+targetNo+"C4").getValue();
    
    //roomID
    params.roomID = targetSheet.getRange("R"+targetNo+"C5").getValue();
    
    //messageID
    params.messageID = targetSheet.getRange("R"+targetNo+"C6").getValue();
    
    //accountID
    params.accountID = targetSheet.getRange("R"+targetNo+"C7").getValue();
    
    //userName
    params.userName = targetSheet.getRange("R"+targetNo+"C8").getValue();
    
    //body
    params.body = targetSheet.getRange("R"+targetNo+"C9").getValue();
    
    //type
    params.type = targetSheet.getRange("R"+targetNo+"C10").getValue();
    
    // 判定したタイプ毎の処理
    switch(params.type){
      // 検索ワードにヒットしないものはスルー
      case 0:
        //返信時間の記入
        targetSheet.getRange("R"+targetNo+"C2").setValue(0);
    
        //返信カウントの記入
        targetSheet.getRange("L1").setValue(targetNo);
        
        continue;
        break;
      // 残り有給取得
      case 9:
        params.result = _getRestNum(params.accountID,params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.result.param)
        break;
      // 休み登録
      case 3:
      case 4:
      case 5:
        params.result = _setRest(params.accountID,params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName);
        break;
      // 休みの事前・事後登録
      case 13:
      case 14:
      case 15:
        params.result = _correctRest(params.accountID,params.body.match(/\d{4}\/\d{2}\/\d{2}/),params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.date);
        break;
        // 特別休暇の事前・事後登録
      case 17:
      case 18:
      case 19:
        params.result = _specialRest(params.accountID,params.body.match(/\d{4}\/\d{2}\/\d{2}/),params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.date);
        break;
      // 休暇の取消処理
      case 20:
        params.result = _resetRest(params.accountID,params.body.match(/\d{4}\/\d{2}\/\d{2}/),params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.date);
        break;
       // 代休取得
      case 21:
        params.result = _compensatrylRest(params.accountID,params.body.match(/\d{4}\/\d{2}\/\d{2}/),String(params.body.match(/→ \d{4}\/\d{2}\/\d{2}/)).match(/\d{4}\/\d{2}\/\d{2}/),params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.dateH).replace('{6}',params.result.dateW);
        break;
      // 勤怠修正
      case 11:
      case 12:
        params.result = _correctAttendance(params.accountID,params.body.match(/\d{4}\/\d{2}\/\d{2}/),params.body.match(/\d{2}\:\d{2}\:\d{2}/),params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.date).replace('{6}',params.result.time);
        break;
      // 勤怠登録
      case 1:
      case 2:
        params.result = _setAttendance(params.accountID,params.time,parseInt(params.type));
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.param);
        break;
      // 予定
      case 16:
        params.result = _getScheduleByDay(params.accountID,new Date(_getScheduleDate(params.body)),params.type);
        params.serif = SERIF_LIST[params.result.code].replace('{1}',params.accountID).replace('{2}',params.roomID).replace('{3}',params.messageID).replace('{4}',params.userName).replace('{5}',params.result.param);
        break;
    }
    
    // 返信が全体部屋かつ、設定がTo All以外
    // もしくは残り有休数取得の場合
    var returnRoomID = params.roomID;
    if(returnRoomID == ROOM_ID_ALL && (vProcessType != 0 || params.type == 9)){
      for(var j in userList){
        if(userList[j].roomID == '')continue;
        if(userList[j].accountID == params.accountID){
          returnRoomID = userList[j].roomID;
          break;
        }
      }
      // この時点でも個別roomIDが取得出来ていない場合はcontinue
      if(returnRoomID == ROOM_ID_ALL)continue;
    }
    
    
    //　返信
    _sendMessage(vAccountToken,returnRoomID,params.serif);
    
    //返信時間の記入
    targetSheet.getRange("R"+targetNo+"C2").setValue(_changeFormatTime(new Date(),0));
    
    //返信カウントの記入
    targetSheet.getRange("L1").setValue(targetNo);
    
  }
};
// ==================================================
// 社員全員のaccountIDを取得
// ==================================================
function _getUserAll(){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetRange = spreadSheet.getRangeByName(RANGE_EMBOSSING_MST);
  
  var userList = new Array();
  
  for(var i = 1 ; i <= targetRange.getLastRow(); i++){
    if(targetRange.getCell(i, 1).getValue() != ''){
      userList[i] = {
        'accountID' : targetRange.getCell(i, 1).getValue(),
        'roomID' : targetRange.getCell(i, 2).getValue()
      };
    }
  }
  
  return userList;
}

// ==================================================
// 社員毎のSheetオブジェクトを取得
// シートの更新処理込み
// ==================================================
function _getSheetByAccount(vAccountID){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetSheet = spreadSheet.getSheetByName(vAccountID);
  
  // アカウントIDセルにアカウントIDを再設定する事で、シート全体に更新をかけて最新の計算結果状態にする
  targetSheet.getRange(CELL_ACCOUNT_ID).setValue(vAccountID);
  
  return targetSheet;
}

// ==================================================
// 有給管理のSheetオブジェクトを取得
// ==================================================
function _getRestSheet(){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  return spreadSheet.getSheetByName(NAME_REST_SHEET);
}

// ==================================================
// 日付フォーマット変換
// type 0:yyyy/MM/dd hh:mm:ss 1:hh:mm:ss
// ==================================================
function _changeFormatTime(vTime,type){
  var time = new Date(vTime);
  switch(type){
    case 0:return time.getFullYear()+'/'+('0'+(time.getMonth()+1)).slice(-2)+'/'+('0'+time.getDate()).slice(-2)+' '+('0'+time.getHours()).slice(-2)+':'+('0'+time.getMinutes()).slice(-2)+':'+('0'+time.getSeconds()).slice(-2);
    case 1:return ('0'+time.getHours()).slice(-2)+':'+('0'+time.getMinutes()).slice(-2)+':'+('0'+time.getSeconds()).slice(-2);
  }
  return new Date();
}

// ==================================================
// 文言から実行するメソッド番号を返す
// ==================================================
function _judgeMethod(body){
  
  for(var i in REGULAR_EXPRESSION){
    result = body.search(REGULAR_EXPRESSION[i]);
    
    if(result != -1){
      return i;
    }
  }
  return '0';
}

// ==================================================
// ChatWorkへメッセージ送信
// ==================================================
function _sendMessage(vAccountToken,vRoomID,vMsg) {
  new ChatWork({token: vAccountToken}).sendMessage({room_id: vRoomID, body: vMsg});
};

// ==================================================
// ChatWorkからメッセージ取得
// ==================================================
function _getMessage(vAccountToken,vRoomID){
  return new ChatWork({token:vAccountToken}).getMessage(vRoomID);
}

// ==================================================
// シートの保護
// ==================================================
function _setProtect(vAccountID){
  
  var targetSheet = _getSheetByAccount(vAccountID);
  
  // 
  var targetRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 既存のシート保護(対象はターゲットのシート内の保護のみ)を全消去(念のため編集権限のチェックも)
  var protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for ( var i in protections ){
    if(protections[i].canEdit()){
      protections[i].remove();
    }
  }

  // 編集可能期間日数以前は編集禁止
  var protectCount = targetRow - EDITABLE_LIMIT;
  var protectRow;
  if(protectCount < 1){
    protectRow = NUM_START_AREA;
  }else{
    protectRow = protectCount + NUM_START_AREA;
  }
  
  // シート保護設定
  targetSheet.getRange('R'+NUM_START_AREA+'C1:R'+protectRow+'C5').protect().setDescription('ID:'+vAccountID+"'s protect");
  
  // スプレッドシートオーナーとスクリプト実行者以外の編集権限を除去
  var protections2 = targetSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for ( var i in protections2 ){
    editorList2 = protections2[i].getEditors();
    for(var j in editorList2){
      protections2[i].removeEditor(editorList2[j]);
    }
  }
}

// ==================================================
// 有休を全員入力しなおす
// ==================================================
function entryRest(){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetRange = spreadSheet.getRangeByName(RANGE_REST_MST);
  
  for(var i = 1 ; i < targetRange.getLastRow(); i++){
    if(targetRange.getCell(i, 1).getValue() == "")continue;
    _setUsedRest(String(targetRange.getCell(i, 1).getValue()));
  }
}

// ==================================================
// 集計
// ==================================================
function outputAttendance(){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetRange = spreadSheet.getRangeByName(RANGE_EMBOSSING_MST);
  
  for(var i = 1 ; i < targetRange.getLastRow(); i++){
    if(targetRange.getCell(i, 1).getValue() == "")continue;
    _aggregate(String(targetRange.getCell(i, 1).getValue()),String(targetRange.getCell(i, 3).getValue()),"2016/04",0);
  }
}

// ==================================================
// 集計(全一覧)
// ==================================================
function outputAttendanceAll(){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetRange = spreadSheet.getRangeByName(RANGE_EMBOSSING_MST);
  
  var totalOut = "CWID,社員名,総労働時間,残業時間,稼働時間,稼働時間(深夜),欠勤時間,稼働時間(土),稼働時間(日),稼働時間(祝),休暇,出勤有無,平日出勤,休日出勤,欠勤\n"
  
  for(var i = 1 ; i < targetRange.getLastRow(); i++){
    if(targetRange.getCell(i, 1).getValue() == "")continue;
    totalOut += _aggregate(String(targetRange.getCell(i, 1).getValue()),String(targetRange.getCell(i, 3).getValue()),"2016/04",1);
  }
   _outputSheetToCsvFile("集計結果",totalOut);
}

// ==================================================
// 正規表現チェック
// ==================================================
function testCheckRegularExpression(){
  Logger.log(_judgeMethod("#特別# 2016/04/25 全休 誕生日"));
}

// ==================================================
// 出勤漏れチェック
// ==================================================
function _checkInputMiss(vAccountToken){
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetDate = new Date();
  targetDate.setDate (targetDate.getDate() - INPUT_MISS_CHECK_DAY);
  
  // 土日祝ならチェックしない
  if(_chechWeekDay(spreadSheet,targetDate) != 0)return;
  
  var targetRange = spreadSheet.getRangeByName(RANGE_EMBOSSING_MST);
  
  var title = "【パルプンテ】打刻漏れ通知";
  
  var result = new Array();
  var toID = "";
  var toName = "";
  var toRoomID = "";
  
  var retStr = "";
  for(var i = 1 ; i < targetRange.getLastRow(); i++){
    toID = targetRange.getCell(i, 1).getValue();
    if(toID == "")continue;
    if(toID == "1746967")continue;
    
    toName = targetRange.getCell(i, 3).getValue();
      result = _checkInputMissByAccount(String(toID),toName,INPUT_MISS_CHECK_DAY);
      
      if(result["flg"]){
        toRoomID = targetRange.getCell(i, 2).getValue();
        if(toRoomID == "")continue;
        _sendMessage(vAccountToken,toRoomID,result["body"]);
      }
  }
}

// ==================================================
// 平日チェック
// ==================================================
function _chechWeekDay(spreadSheet,vDate){
  
  var targetDate = vDate;
  var targetMonth = targetDate.getMonth()+1;
  
  // 経過日数を求める(＝シートの取得位置)
  var courseDate = 0;
  for(var i = 1;i < targetMonth ;i++){
    courseDate += new Date(parseInt(targetDate.getFullYear(), 10), parseInt(i, 10), 0).getDate();
  }
  
  courseDate += targetDate.getDate();
  
  var targetSheet = spreadSheet.getSheetByName("暦");
  
  return targetSheet.getRange(courseDate, 2).getValue();
  
}

// ==================================================
// パルプンテからのメール送信
// ==================================================
function _sendMailByParupunte(vTo,vTitle,vBody){
  if(ENABLE_GMAIL){
    GmailApp.sendEmail(vTo, vTitle, vBody);
  }
}

// ==================================================
// 終日予定の登録
// ==================================================
function _setScheduleAllDay(vTitle,vDate){
  if(ENABLE_GOOGLE_CALENDER){
    var cal = CalendarApp.getCalendarById(GOOGLE_CALENDER_ID);
    var event = cal.createAllDayEvent(vTitle,new Date(vDate)); 
    return event.getId();
  }
}