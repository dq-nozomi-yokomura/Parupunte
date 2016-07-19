// ==================================================
// 有給休暇残り日数取得
// account_id：アカウントID
// type 
// ==================================================
function _getRestNum(vAccountID,vType){
  var targetSheet = _getSheetByAccount(vAccountID);
  var num = targetSheet.getRange(CELL_REST).getValue();// 残り日数セルから取得
  var roomID = targetSheet.getRange(CELL_ROOM_ID).getValue();// 管理ユーザとの個別チャットルーム番号セルから取得
  
  if(parseInt(num)>0)return{'code':'2'+('0'+vType).slice(-2),'param':num,'room':roomID};
  else return{'code':'4'+('0'+type).slice(-2),'param':'','room':roomID};
}

// ==================================================
// 有給休暇使用数入力
// account_id：アカウントID
// type 
// ==================================================
function _setUsedRest(vAccountID){
  // 記入位置を取得
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  var targetRange = spreadSheet.getRangeByName('RestMst');
  var restIndex = 0;
  
  for(var i = 1 ; i < targetRange.getLastRow(); i++){
    if(targetRange.getCell(i, 1).getValue() == vAccountID){
      restIndex = i;
      break;
    }
  }
  // 記入位置が見つからなかった場合は何もせずに終了
  if(restIndex == 0)return;
  
  var targetSheet = _getSheetByAccount(vAccountID);
  var usedRestNum = targetSheet.getRange(CELL_USED_REST_NUM).getValue();// 消化有給数取得
  var restSheet = _getRestSheet();//有給管理シート取得
  
  // 有給管理はヘッダがあるので1行プラス
  // 消化数は4列目
  restSheet.getRange(restIndex+1,4).setValue(usedRestNum);
}

// ==================================================
// 出退勤の記入
// sheetName : シート名
// setTime：設定する時間
// type 1:出勤 2:退勤
// ==================================================
function _setAttendance(vAccountID,setTime,type){
  var targetSheet = _getSheetByAccount(vAccountID);
  
  // 返信ルームID
  var roomID = targetSheet.getRange(CELL_ROOM_ID).getValue();// 管理ユーザとの個別チャットルーム番号セルから取得
  
  // 再入力可能フラグのセルをチェック(5分以上の同一勤怠エラー)
  if(targetSheet.getRange(1, 1+type).getValue() == "NG"){return {'code':'4'+('0'+type).slice(-2),'param':'','room':roomID};}
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 出勤が2列目、退勤が3列目
  var setColumnNum = parseInt(type) + 1;
  
  // 入力
  targetSheet.getRange(setRowNum,setColumnNum).setValue(setTime);

  // 正常終了
  return {'code':'2'+('0'+type).slice(-2),'param':_changeFormatTime(setTime,1),'room':roomID};
}

// ==================================================
// 休みの記入
// ==================================================
function _setRest(vAccountID,type){
  var targetSheet = _getSheetByAccount(vAccountID);
  
  var restTime;
  var setParam;
  switch(type){
    case 3:
      restTime = 1;
      setParam = '全休';
      break;
    case 4:
      restTime = 0.5;
      setParam = '午前休';
      break;
    case 5:
      restTime = 0.5;
      setParam = '午後休';
      break;
  }
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 入力
  targetSheet.getRange(setRowNum,4).setValue(setParam);
  
  // 使用有休数の入力
  _setUsedRest(vAccountID);
  
  // Googleカレンダー登録用の名前取得
  var userName = targetSheet.getRange(CELL_USER_NAME).getValue();
  var pos = userName.search(" ");
  if(pos != -1){
    userName = userName.substring(0,pos);
  }
  
  var targetDate = targetSheet.getRange(CELL_TODAY_DATE).getValue();
    
  // Googleカレンダーに登録
  _setScheduleAllDay("【"+userName+"】"+setParam,targetDate);
  
  // 正常終了
  return {'code':'2'+('0'+type).slice(-2)};
}

// ==================================================
// 休みの修正
// ==================================================
function _correctRest(vAccountID,vDate,vType){
  var targetSheet = _getSheetByAccount(vAccountID);
  
  var restTime;
  var setParam;
  switch(vType%10){
    case 3:
      restTime = 1;
      setParam = '全休';
      break;
    case 4:
      restTime = 0.5;
      setParam = '午前休';
      break;
    case 5:
      restTime = 0.5;
      setParam = '午後休';
      break;
  }
  
  // 修正日の入力を行う
  targetSheet.getRange(CELL_CHANGE_DATE).setValue(vDate);
  
  // 編集禁止チェック用に本日のインデックスを取得
  var checkRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_CHANGE_INDEX).getValue();
  
  if(checkRow - targetRow >= EDITABLE_LIMIT){
    return {'code':'4'+('0'+vType).slice(-2),'date':vDate};
  }
  
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 入力
  targetSheet.getRange(setRowNum,4).setValue(setParam);
  
  // 使用有休数の入力
  _setUsedRest(vAccountID);
  
  // Googleカレンダー登録用の名前取得
  var userName = targetSheet.getRange(CELL_USER_NAME).getValue();
  var pos = userName.search(" ");
  if(pos != -1){
    userName = userName.substring(0,pos);
  }
    
  // Googleカレンダーに登録
  _setScheduleAllDay("【"+userName+"】"+setParam,vDate);
  
  // 正常終了
  return {'code':'2'+('0'+vType).slice(-2),'date':vDate};
}

// ==================================================
// 特別休暇の登録
// ==================================================
function _specialRest(vAccountID,vDate,vType){
  var targetSheet = _getSheetByAccount(vAccountID);
  
  var restTime;
  var setParam;
  switch(vType){
    case 17:
      restTime = 1;
      setParam = '全特休';
      break;
    case 18:
      restTime = 0.5;
      setParam = '午前特休';
      break;
    case 19:
      restTime = 0.5;
      setParam = '午後特休';
      break;
  }
  
  // 登録日の入力を行う
  targetSheet.getRange(CELL_CHANGE_DATE).setValue(vDate);
  
    // 編集禁止チェック用に本日のインデックスを取得
  var checkRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_CHANGE_INDEX).getValue();
  
  if(checkRow - targetRow >= EDITABLE_LIMIT){
    return {'code':'4'+('0'+vType).slice(-2),'date':vDate};
  }
  
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 入力
  targetSheet.getRange(setRowNum,4).setValue(setParam);
  
  // Googleカレンダー登録用の名前取得
  var userName = targetSheet.getRange(CELL_USER_NAME).getValue();
  var pos = userName.search(" ");
  if(pos != -1){
    userName = userName.substring(0,pos);
  }
    
  // Googleカレンダーに登録
  _setScheduleAllDay("【"+userName+"】"+setParam,vDate);
  
  // 正常終了
  return {'code':'2'+('0'+vType).slice(-2),'date':vDate};
}

// ==================================================
// 休みの取り消し
// ==================================================
function _resetRest(vAccountID,vDate,vType){
  var targetSheet = _getSheetByAccount(vAccountID);
    
  // 修正日の入力を行う
  targetSheet.getRange(CELL_CHANGE_DATE).setValue(vDate);
  
  // 編集禁止チェック用に本日のインデックスを取得
  var checkRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_CHANGE_INDEX).getValue();
  
  if(checkRow - targetRow >= EDITABLE_LIMIT){
    return {'code':'4'+('0'+vType).slice(-2),'date':vDate};
  }
  
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 入力
  targetSheet.getRange(setRowNum,4).setValue("");
  
  // 使用有休数の入力
  _setUsedRest(vAccountID);
  
  // 正常終了
  return {'code':'2'+('0'+vType).slice(-2),'date':vDate};
}


// ==================================================
// 代休の登録
// ==================================================
function _compensatrylRest(vAccountID,vDateHoliday,vDateWeekday,vType){
  var targetSheet = _getSheetByAccount(vAccountID);
  
  // 登録日の入力を行う
  targetSheet.getRange(CELL_CHANGE_DATE).setValue(vDateHoliday);
  
  // 編集禁止チェック用に本日のインデックスを取得
  var checkRowH = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 出退勤の登録エリアは6行目から
  var targetRowH = targetSheet.getRange(CELL_CHANGE_INDEX).getValue();
  
  if(checkRowH - targetRowH >= EDITABLE_LIMIT){
    return {'code':'4'+('0'+vType).slice(-2),'dateH':vDateHoliday,'dateW':vDateWeekday};
  }
  
  var setRowNumH = targetRowH + NUM_START_AREA;
  
  
  
  // 登録日の入力を行う
  targetSheet.getRange(CELL_CHANGE_DATE).setValue(vDateWeekday);
  
    // 編集禁止チェック用に本日のインデックスを取得
  var checkRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_CHANGE_INDEX).getValue();
  
  if(checkRow - targetRow >= EDITABLE_LIMIT){
    return {'code':'4'+('0'+vType).slice(-2),'dateH':vDateHoliday,'dateW':vDateWeekday};
  }
  
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 入力
  targetSheet.getRange(setRowNumH,4).setValue("代休取得済");
  
  // 入力
  targetSheet.getRange(setRowNum,4).setValue("代休");
  
  // Googleカレンダー登録用の名前取得
  var userName = targetSheet.getRange(CELL_USER_NAME).getValue();
  var pos = userName.search(" ");
  if(pos != -1){
    userName = userName.substring(0,pos);
  }
    
  // Googleカレンダーに登録
  _setScheduleAllDay("【"+userName+"】代休",vDateWeekday);
  
  // 正常終了
  return {'code':'2'+('0'+vType).slice(-2),'dateH':vDateHoliday,'dateW':vDateWeekday};
}

// ==================================================
// 勤怠修正
// account_id : アカウントID(シート名に設定している)
// setTime：設定する時間
// type 11:出勤 12:退勤
// ==================================================
function _correctAttendance(ｖAccountID,vDate,vTime,vType){
  
  var targetSheet = _getSheetByAccount(ｖAccountID);
  
  // 修正日の入力を行う
  targetSheet.getRange(CELL_CHANGE_DATE).setValue(vDate);
  
  // 編集禁止チェック用に本日のインデックスを取得
  var checkRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_CHANGE_INDEX).getValue();
  
  if(checkRow - targetRow >= EDITABLE_LIMIT){
    return {'code':'4'+('0'+vType).slice(-2),'date':'','time':''};
  }
  
  var setRowNum = targetRow + NUM_START_AREA;
  
  // 出勤が2列目、退勤が3列目
  var setColumnNum = vType%10 + 1;
  
  // 入力
  targetSheet.getRange(setRowNum,setColumnNum).setValue(vDate+' '+vTime);
  targetSheet.getRange(setRowNum,setColumnNum).setNumberFormat('h":"mm":"ss');

  // 正常終了
  return {'code':'2'+('0'+vType).slice(-2),'date':vDate,'time':vTime};
}

// ==================================================
// 入力漏れのチェック
// ==================================================
function _checkInputMissByAccount(vAccountID,vName,vDayCount){
  var targetSheet = _getSheetByAccount(vAccountID);
  
  // 出退勤の登録エリアは6行目から
  var targetRow = targetSheet.getRange(CELL_TODAY_INDEX).getValue();
  var setRowNum = targetRow + NUM_START_AREA - vDayCount;
  
  // 
  var attendanceIN = targetSheet.getRange(setRowNum,2).getValue();
  var attendanceOUT = targetSheet.getRange(setRowNum,3).getValue();
  var rest = targetSheet.getRange(setRowNum,4).getValue();
  
  var messageStart = "[To:"+vAccountID+"] ドジっ子な"+vName+"さん\n\n";
  var messageEnd = MESSAGE_LIST[‘err002’];
  
  var targetDate = new Date(targetSheet.getRange(CELL_TODAY_DATE).getValue());
  targetDate.setDate (targetDate.getDate() - vDayCount);
  targetDate = ('0'+(targetDate.getMonth()+1)).slice(-2)+'月'+('0'+targetDate.getDate()).slice(-2)+'日';
  
  if(rest == "全休"　|| rest == "全特休" || rest == "代休"){
    return "";
  }else{
    if(attendanceIN == "" && attendanceOUT =="" ){
      return {flg:true,body:messageStart+targetDate+":出勤と退勤がどちらも入力されていないよ。\n"+messageEnd};
    }else if(attendanceIN == "" ){
      return {flg:true,body:messageStart+targetDate+":出勤が入力されてないよ。\n"+messageEnd};
    }else if(attendanceOUT =="" ){
      return {flg:true,body:messageStart+targetDate+":退勤が入力されてないよ。\n"+messageEnd};
    }else{
      return {flg:false};
    }
  }
}