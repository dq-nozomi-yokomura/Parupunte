// ==================================================
// csv出力
// ==================================================
function _outputSheetToCsvFile(vFileName,vCsvString) {
  
  var contentType = "text/csv";
  var charSet = "Shift_JIS";

  // Shift_JISなBlobに変換
  var blob = Utilities.newBlob("", contentType, vFileName).setDataFromString(vCsvString, charSet);

  // Blobをファイルに出力
  var fileObj = DriveApp.createFile(blob);
  
  // 移動先のフォルダオブジェクトを取得
  var hoge_folder = DriveApp.getFolderById(DIR_AGGREGATE_FILE);
  
  // 移動先へadd、ルートフォルダ(たぶんマイドライブ)からremove、のコンボで移動
  hoge_folder.addFile(fileObj);
  DriveApp.getRootFolder().removeFile(fileObj);
  
}

// ==================================================
// 集計作業
// ==================================================
function _aggregate(vAccountID,vUserName,vTargetMonth,vType){
  
  var spreadSheet = SpreadsheetApp.openByUrl(WORKBOOK_URL);
  try{
    var targetSheet = spreadSheet.getSheetByName(vAccountID);
 
    var workRange = targetSheet.getRange(_getWorkRange(vTargetMonth));
  }catch(e){return;}
  
  var Params = function(){
    this.date;
    this.start;
    this.end;
    this.workTimeMorning;
    this.workTimeAfternoon;
    this.workTimeNight;
    this.absenceTimeMorning;
    this.absenceTimeAfternoon;
    this.type;
    this.state;
    this.attendance;
  }
  
  var targetList = new Array();
  
  var limit = workRange.getHeight();
  var params;
  
  // 土日祝のデータを取得
  var restDateList = _getRestDate(spreadSheet,vTargetMonth);

  for(var i = 1 ; i <= limit ; i++){
  
    params = new Params();

    params.date = workRange.getCell(i, 1).getValue();
    params.type = restDateList[i-1];
    params.state = workRange.getCell(i, 4).getValue();
    
    // 未入力があったら非出社扱い
    if(workRange.getCell(i, 2).getValue() == "" || workRange.getCell(i, 3).getValue() == ""){
      params.attendance = 0;
      // 平日だった場合
      if(params.type == 0 && params.state != "全休"　&& params.state != "全特休"　&& params.state != "代休"){
        params.attendance = 2;
      }
      targetList.push(params);
      continue;
    }else{
      params.attendance = 1;
    }
    
    params.start = new Date(workRange.getCell(i, 2).getValue());
    params.end = new Date(workRange.getCell(i, 3).getValue());
    
    //Logger.log(_getDistance(params.start,params.end));
    
    // 出社時間より退社時間の方が先ならエラー
    if(params.start > params.end){
      params.attendance = 3;
      targetList.push(params);
      continue;
    }
    
    //--------------------------------------------------
    //①総労働時間
    //・時間　ただし13:00〜14:00は除く
    //②休日
    //・時間(土・日・祝)　ただし13:00～14:00は除く
    //③深夜
    //・22:00〜5:00(平日土日祝まとめる)
    //④欠勤時間
    //・10:00〜13:00、14:00〜19:00から欠けている時間
    //--------------------------------------------------
    
    var targetDate = new Date(params.date);
    targetDate = targetDate.getFullYear()+'/'+('0'+(targetDate.getMonth()+1)).slice(-2)+'/'+('0'+targetDate.getDate()).slice(-2);
    var startTime = new Date(targetDate+' 10:00:00');
    var restTimeStart = new Date(targetDate+' 13:00:00');
    var restTimeEnd = new Date(targetDate+' 14:00:00');
    var endTime = new Date(targetDate+' 19:00:00');
    var nightTime = new Date(targetDate+' 22:00:00');
    
    // 〜13:00に出社
    if(params.start <= restTimeStart){
      
      // 欠勤時間(10:00〜の出社)
      if(startTime < params.start){
        params.absenceTimeMorning = _getDistance(startTime,params.start);
      }else{
        params.absenceTimeMorning = 0;
      }
      
      // 〜13:00に退社
      if(params.end <= restTimeStart){
        
        // 欠勤時間(〜10:00の退社)
        if(params.end < startTime){
          params.absenceTimeMorning += _getDistance(startTime,restTimeStart);
        }else{
          params.absenceTimeMorning += _getDistance(params.end,restTimeStart);
        }
        
        params.absenceTimeAfternoon = _getDistance(restTimeEnd,endTime);
        
        params.workTimeMorning = _getDistance(params.start,params.end);
        params.workTimeAfternoon = 0;
        params.workTimeNight = 0;
      }
      
      // 13:00〜14:00に退社
      if(restTimeStart < params.end && params.end < restTimeEnd){
        
        // 午後の時間がフルで欠勤時間
        params.absenceTimeAfternoon=_getDistance(startTime,restTimeStart);
        
        params.workTimeMorning = _getDistance(params.start,restTimeStart);
        params.workTimeAfternoon = 0;
        params.workTimeNight = 0;
      }
      
      // 14:00〜に退社
      if(restTimeEnd <= params.end){
        
        params.workTimeMorning = _getDistance(params.start,restTimeStart);
        
        // 〜22:00に退社
        if(params.end < nightTime){
          
          // 欠勤時間(10:00〜の出社)
          if(params.end < endTime){
            params.absenceTimeAfternoon = _getDistance(params.end,endTime);
          }else{
            params.absenceTimeAfternoon = 0;
          }
          
          params.workTimeAfternoon = _getDistance(restTimeEnd,params.end);
          params.workTimeNight = 0;
        } //　22:00〜に退社
        else{
          params.absenceTimeAfternoon = 0;
          params.workTimeAfternoon = _getDistance(restTimeEnd,nightTime);
          params.workTimeNight = _getDistance(nightTime,params.end);
        }
        
      }
      
    } // 13:00〜14:00に出社
    else if(restTimeStart < params.start && params.start < restTimeEnd){
      
      // 午前中の稼働時間は無し
      params.workTimeMorning = 0;
      
      // 午前中がフルで欠勤時間
      params.absenceTimeMorning = _getDistance(startTime,restTimeStart);
      
      // 〜22:00に退社
      if(params.end < nightTime){
        
        // 欠勤時間(〜19:00の退社)
        if(params.end < endTime){
          params.absenceTimeAfternoon = _getDistance(params.end,endTime);
        }else{
          params.absenceTimeAfternoon = 0;
        }
        
        params.workTimeAfternoon = _getDistance(restTimeEnd,params.end);
        params.workTimeNight = 0;
      } //　22:00〜に退社
      else{
        params.absenceTimeAfternoon = 0;
        params.workTimeAfternoon = _getDistance(restTimeEnd,nightTime);
        params.workTimeNight = _getDistance(nightTime,params.end);
      }
      
    } // 14:00〜に出社
    else if(restTimeEnd <= params.start){
      
      // 午前中の稼働時間は無し
      params.workTimeMorning = 0;
      
      // 午前中がフルで欠勤時間
      params.absenceTimeMorning = _getDistance(startTime,restTimeStart);
      params.absenceTimeAfternoon= 0;
      // 欠勤時間(14:00〜の出社)
      
      if(params.start < endTime){
        params.absenceTimeAfternoon = _getDistance(restTimeEnd,params.start);
      }else{
        params.absenceTimeAfternoon = _getDistance(restTimeEnd,endTime);
      }
      
      // 〜22:00に退社
      if(params.end < nightTime){
        
        // 欠勤時間(〜19:00の退社)
        if(params.end < endTime){
          params.absenceTimeAfternoon += _getDistance(params.end,endTime);
        }else{
          params.absenceTimeAfternoon += 0;
        }
        
        params.workTimeAfternoon = _getDistance(params.start,params.end);
        params.workTimeNight = 0;
      } //　22:00〜に退社
      else{
        
        // 欠勤時間(〜19:00の退社)
        if(params.end < endTime){
          params.absenceTimeAfternoon += _getDistance(params.end,endTime);
        }else{
          params.absenceTimeAfternoon += 0;
        }
        
        if(params.start < nightTime){
          params.workTimeAfternoon = _getDistance(params.start,nightTime);
          params.workTimeNight = _getDistance(nightTime,params.end);
        }else{
          params.workTimeAfternoon = 0;
          params.workTimeNight = _getDistance(params.start,params.end);
        }
      }
    }
    targetList.push(params);
  
  }
  
  
  //--------------------------------------------------
  // 総合集計
  //--------------------------------------------------
  
  var summaryTotalTime = 0;
  var summaryTotalOvertime = 0;
  var summaryWorkTimeM = 0;
  var summaryWorkTimeA = 0;
  var summaryWorkTimeN = 0;
  var summaryAbsenceTimeM = 0;
  var summaryAbsenceTimeA = 0;
  var summaryWorkTimeSatM = 0;
  var summaryWorkTimeSatA = 0;
  var summaryWorkTimeSunM = 0;
  var summaryWorkTimeSunA = 0;
  var summaryWorkTimePubM = 0;
  var summaryWorkTimePubA = 0;
  
  var summaryWorkDateN = 0;
  var summaryWorkDateR = 0;
  var summaryAbsence = 0;
  var summaryUsedRest = 0;
  var summaryUsedRestS = 0;
  
  var strOutput = "日付,始業時間,終業時間,総労働時間,残業時間,稼働時間,稼働時間(深夜),欠勤時間,稼働時間(土),稼働時間(日),稼働時間(祝),平日/休日,休暇,出勤有無,平日出勤,休日出勤,欠勤\n"
  var totalTime = 0;
  var overTime = 0;
  
  for(var j in targetList){
    totalTime = 0;
    overTime = 0;
    
    // 出勤無しの日
    if(targetList[j].attendance == 0){
      strOutput += 
        ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        targetList[j].type +","+
        targetList[j].state +","+
        targetList[j].attendance +",,,\n";
      
      if(targetList[j].state == "全休"){
            summaryUsedRest += 1;
       }else if(targetList[j].state == "全特休"){
            summaryUsedRestS += 1;
       }
    // エラー
    }else if(targetList[j].attendance == 3){
      strOutput += 
      ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
        _changeFormatTime(targetList[j].start,1) +","+
        _changeFormatTime(targetList[j].end,1) +","+
        ","+
        ","+
        "error,"+
        "error,"+
        ","+
        ","+
        ","+
        targetList[j].type +","+
        targetList[j].state +","+
        targetList[j].attendance +",,,\n";
    }else if(targetList[j].attendance == 2){
      strOutput += 
      ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        ","+
        targetList[j].type +","+
        targetList[j].state +","+
        targetList[j].attendance +",,,1\n";
      summaryAbsence++;
    }else{
      // 平日の出勤の場合
      if(targetList[j].type == 0　|| targetList[j].state == "代休取得済"){
        if(targetList[j].state == "全休" || targetList[j].state == "全特休" || targetList[j].state == "代休"){
          strOutput += 
            ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
            ","+
            ","+
            ","+
            ","+
            ","+
            ","+
            ","+
            ","+
            ","+
            ","+
            targetList[j].type +","+
            targetList[j].state +","+
            targetList[j].attendance +",,,\n";
          
          if(targetList[j].state == "全休"){
            summaryUsedRest += 1;
          }else if(targetList[j].state == "全特休"){
            summaryUsedRestS += 1;
          }
        }else if(targetList[j].state == "午前休"　|| targetList[j].state == "午前特休"　){
          
          totalTime = targetList[j].workTimeAfternoon + targetList[j].workTimeNight;
          overTime = _getOvertime(targetList[j].workTimeAfternoon + targetList[j].workTimeNight );
          
          strOutput += 
            ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
            _changeFormatTime(targetList[j].start,1) +","+
            _changeFormatTime(targetList[j].end,1) +","+
            _getHMS(totalTime,1) +","+
            _getHMS(overTime,1) +","+
            _getHMS(targetList[j].workTimeAfternoon,1) +","+
            _getHMS(targetList[j].workTimeNight,1) +","+
            _getHMS(targetList[j].absenceTimeAfternoon,1) +","+
            ","+
            ","+
            ","+
            targetList[j].type +","+
            targetList[j].state +","+
            targetList[j].attendance +",1,,\n";
      
          summaryTotalTime += totalTime;
          summaryTotalOvertime += overTime;
          summaryWorkTimeM += 0;
          summaryWorkTimeA += targetList[j].workTimeAfternoon;
          summaryWorkTimeN += targetList[j].workTimeNight;
          summaryAbsenceTimeM += 0;
          summaryAbsenceTimeA += targetList[j].absenceTimeAfternoon;
          summaryWorkDateN++;
          if(targetList[j].state == "午前休"){
            summaryUsedRest += 0.5;
          }else if(targetList[j].state == "午前特休"){
            summaryUsedRestS += 0.5;
          }
        }else if(targetList[j].state == "午後休" || targetList[j].state == "午後特休"){
          
          totalTime = targetList[j].workTimeMorning;
          overTime = _getOvertime(targetList[j].workTimeMorning);
          
          strOutput += 
            ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
            _changeFormatTime(targetList[j].start,1) +","+
            _changeFormatTime(targetList[j].end,1) +","+
            _getHMS(totalTime,1) +","+
            _getHMS(overTime,1) +","+
            _getHMS(targetList[j].workTimeMorning,1) +","+
            _getHMS(0) +","+
            _getHMS(targetList[j].absenceTimeMorning,1) +","+
            ","+
            ","+
            ","+
            targetList[j].type +","+
            targetList[j].state +","+
            targetList[j].attendance +",1,,\n";
      
          summaryTotalTime += totalTime;
          summaryTotalOvertime += overTime;
          summaryWorkTimeM += targetList[j].workTimeMorning;
          summaryWorkTimeA += 0;
          summaryWorkTimeN += 0;
          summaryAbsenceTimeM += targetList[j].absenceTimeMorning;
          summaryAbsenceTimeA += 0;
          summaryWorkDateN++;
          if(targetList[j].state == "午後休"){
            summaryUsedRest += 0.5;
          }else if(targetList[j].state == "午後特休"){
            summaryUsedRestS += 0.5;
          }
        }else{
          
          totalTime = targetList[j].workTimeMorning + targetList[j].workTimeAfternoon + targetList[j].workTimeNight;
          overTime = _getOvertime(targetList[j].workTimeMorning + targetList[j].workTimeAfternoon + targetList[j].workTimeNight);
          
          strOutput += 
            ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
            _changeFormatTime(targetList[j].start,1) +","+
            _changeFormatTime(targetList[j].end,1) +","+
            _getHMS(totalTime,1) +","+
            _getHMS(overTime,1) +","+
            _getHMS(targetList[j].workTimeMorning + targetList[j].workTimeAfternoon,1) +","+
            _getHMS(targetList[j].workTimeNight,1) +","+
            _getHMS(targetList[j].absenceTimeMorning + targetList[j].absenceTimeAfternoon,1) +","+
            ","+
            ","+
            ","+
            targetList[j].type +","+
            targetList[j].state +","+
            targetList[j].attendance +",1,,\n";
      
          summaryTotalTime += totalTime;
          summaryTotalOvertime += overTime;
          summaryWorkTimeM += targetList[j].workTimeMorning;
          summaryWorkTimeA += targetList[j].workTimeAfternoon;
          summaryWorkTimeN += targetList[j].workTimeNight;
          summaryAbsenceTimeM += targetList[j].absenceTimeMorning;
          summaryAbsenceTimeA += targetList[j].absenceTimeAfternoon;
          summaryWorkDateN++;
        }
      // 休日出勤
      }else{
        
        totalTime = targetList[j].workTimeMorning + targetList[j].workTimeAfternoon + targetList[j].workTimeNight;
        
        if(targetList[j].type == 1){
          strOutput += 
          ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
          _changeFormatTime(targetList[j].start,1) +","+
          _changeFormatTime(targetList[j].end,1) +","+
          _getHMS(totalTime,1) +","+
          "00:00:00,"+
          ","+
          _getHMS(targetList[j].workTimeNight,1) +","+
          ","+
          _getHMS(targetList[j].workTimeMorning + targetList[j].workTimeAfternoon,1) +","+
          ","+
          ","+
          targetList[j].type +","+
          targetList[j].state +","+
          targetList[j].attendance +",,1,\n";
        }else if(targetList[j].type == 2){
          strOutput += 
          ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
          _changeFormatTime(targetList[j].start,1) +","+
          _changeFormatTime(targetList[j].end,1) +","+
          _getHMS(totalTime,1) +","+
          "00:00:00,"+
          ","+
          _getHMS(targetList[j].workTimeNight,1) +","+
          ","+
          ","+
          _getHMS(targetList[j].workTimeMorning + targetList[j].workTimeAfternoon,1) +","+
          ","+
          targetList[j].type +","+
          targetList[j].state +","+
          targetList[j].attendance +",,1,\n";
        }else if(targetList[j].type == 3){
          strOutput += 
          ('0'+(targetList[j].date.getMonth()+1)).slice(-2)+'/'+('0'+targetList[j].date.getDate()).slice(-2) +","+
          _changeFormatTime(targetList[j].start,1) +","+
          _changeFormatTime(targetList[j].end,1) +","+
          _getHMS(totalTime,1) +","+
          "00:00:00,"+
          ","+
          _getHMS(targetList[j].workTimeNight,1) +","+
          ","+
          ","+
          ","+
          _getHMS(targetList[j].workTimeMorning + targetList[j].workTimeAfternoon,1) +","+
          targetList[j].type +","+
          targetList[j].state +","+
          targetList[j].attendance +",,1,\n";
        } 
          
        summaryTotalTime += totalTime;
        if(targetList[j].type == 1){
          summaryWorkTimeSatM += targetList[j].workTimeMorning;
          summaryWorkTimeSatA += targetList[j].workTimeAfternoon;
        }else if(targetList[j].type == 2){
          summaryWorkTimeSunM += targetList[j].workTimeMorning;
          summaryWorkTimeSunA += targetList[j].workTimeAfternoon;
        }else if(targetList[j].type == 3){
          summaryWorkTimePubM += targetList[j].workTimeMorning;
          summaryWorkTimePubA += targetList[j].workTimeAfternoon;
        }
        
        summaryWorkTimeN += targetList[j].workTimeNight;
        summaryWorkDateR++;
      }
    }
  }
  strOutput += "計,-,-,"+
    _getHMS(summaryTotalTime)+","+
    _getHMS(summaryTotalOvertime)+","+
    _getHMS(summaryWorkTimeM + summaryWorkTimeA)+","+
    _getHMS(summaryWorkTimeN)+","+
    _getHMS(summaryAbsenceTimeM + summaryAbsenceTimeA)+","+
    _getHMS(summaryWorkTimeSatM + summaryWorkTimeSatA)+","+
    _getHMS(summaryWorkTimeSunM + summaryWorkTimeSunA)+","+
    _getHMS(summaryWorkTimePubM + summaryWorkTimePubA)+",-,"+
    summaryUsedRest+","+
    (summaryWorkDateN+summaryWorkDateR)+","+
    summaryWorkDateN+","+
    summaryWorkDateR+","+
    summaryAbsence+"\n";
  
  if(vType == 0){
    _outputSheetToCsvFile(vAccountID+"_"+vUserName+_getTitleMonth(vTargetMonth),strOutput);
  }else{
    var strOutput2 = vAccountID+","+vUserName+","+
    _getHMS(summaryTotalTime)+","+
    _getHMS(summaryTotalOvertime)+","+
    _getHMS(summaryWorkTimeM + summaryWorkTimeA)+","+
    _getHMS(summaryWorkTimeN)+","+
    _getHMS(summaryAbsenceTimeM + summaryAbsenceTimeA)+","+
    _getHMS(summaryWorkTimeSatM + summaryWorkTimeSatA)+","+
    _getHMS(summaryWorkTimeSunM + summaryWorkTimeSunA)+","+
    _getHMS(summaryWorkTimePubM + summaryWorkTimePubA)+","+
    summaryUsedRest+","+
    summaryUsedRestS+","+
    (summaryWorkDateN+summaryWorkDateR)+","+
    summaryWorkDateN+","+
    summaryWorkDateR+","+
    summaryAbsence+"\n";
    
    if(vType == 2){
      _outputSheetToCsvFile(vAccountID+"_"+vUserName+_getTitleMonth(vTargetMonth),strOutput);
    }
    
    return strOutput2;
  }
}

// ==================================================
// 
// ==================================================
function _getTitleMonth(vDate){
  var targetDate = new Date(vDate+"/01");
  return "_"+targetDate.getFullYear()+"_"+(targetDate.getMonth()+1);
}

// ==================================================
// 秒データを時分秒表記になおす
// 3700(s) → 1:01:40
// ==================================================
function _getHMS(vSecond){
  var hour = Math.floor(vSecond / 3600);
  var minutes = Math.floor((vSecond % 3600) / 60);
  var second = vSecond % 60;
  return hour+":"+minutes+":"+second;
}

// ==================================================
// 残業時間を秒で返す
// ==================================================
function _getOvertime(vSecond){
  var hour = Math.floor(vSecond / 3600);
  
  if(hour >= 8){
    return parseInt(vSecond - 8*3600);
  }else{
    return 0;
  }
}

// ==================================================
// 時間差をhour単位(小数点含む)で返す
// ==================================================
function _getDistance(src,dst){
  var SECOND_MILLISECOND = 1000;
  var MINUTE_MILLISECOND = 60 * SECOND_MILLISECOND;
  var HOUR_MILLISECOND = 60 * MINUTE_MILLISECOND;
  var DAY_MILLISECOND = 24 * HOUR_MILLISECOND;
  var WEEK_MILLISECOND = 7 * DAY_MILLISECOND;
  var YEAR_MILLISECOND = 365 * DAY_MILLISECOND;

  var deltaMillsecond = dst.getTime() - src.getTime();
  return deltaMillsecond / SECOND_MILLISECOND;

}

// ==================================================
// R1C1形式で作業範囲を返す
// ==================================================
function _getWorkRange(vDate){

  var targetDate = new Date(vDate+"/01");
  var targetMonth = targetDate.getMonth()+1;
  
  // 経過日数を求める(＝シートの取得位置)
  var courseDate = 0;
  for(var i = 1;i < targetMonth ;i++){
    courseDate += new Date(parseInt(targetDate.getFullYear(), 10), parseInt(i, 10), 0).getDate();
  }
  
  // 対象月の日数を求める
  var dateNum = new Date(parseInt(targetDate.getFullYear(), 10), parseInt(targetMonth, 10), 0).getDate();

  // ヘッダまでが5行なので6行目が開始位置
  var start = 6 + courseDate;
  var ret = "R"+start+"C1:R"+parseInt(start+dateNum-1)+"C4";
  
  return ret;
}

// ==================================================
// 土日祝のフラグ配列を返す
// ==================================================
function _getRestDate(spreadSheet,vDate){
  var targetSheet = spreadSheet.getSheetByName("暦");
  
  var targetDate = new Date(vDate+"/01");
  var targetMonth = targetDate.getMonth()+1;
  
  // 経過日数を求める(＝シートの取得位置)
  var courseDate = 0;
  for(var i = 1;i < targetMonth ;i++){
    courseDate += new Date(parseInt(targetDate.getFullYear(), 10), parseInt(i, 10), 0).getDate()
  }
  
  // 対象月の日数を求める
  var dateNum = new Date(parseInt(targetDate.getFullYear(), 10), parseInt(targetMonth, 10), 0).getDate();
  
  // ヘッダ無しなので1行目が開始位置
  var start = 1 + courseDate;
  var ret = "R"+start+"C1:R"+parseInt(start+dateNum-1)+"C2";
  
  var workRange = targetSheet.getRange(ret);
  
  var retList = new Array();
  
  for(var i = 1 ; i <= workRange.getHeight() ; i++ ){
    retList.push(workRange.getCell(i,2).getValue());
  }
  
  return retList;
}

// ==================================================
// 小数点n位までを残す関数
// number=対象の数値
// n=残したい小数点以下の桁数
// ==================================================
function floatFormat( number, n ) {
	var _pow = Math.pow( 10 , n ) ;

	return Math.floor( number * _pow ) / _pow ;
}
