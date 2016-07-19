// attendance book
var WORKBOOK_URL = "https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxx/edit#gid=0";

// log book
var LOGBOOK_URL = "https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxx/edit#gid=0";

// aggregate file's directory
var DIR_AGGREGATE_FILE = 'xxxxxxxxxxxxxxxxxxxxxxxxx';

// google calender's account id
var GOOGLE_CALENDER_ID = 'xxxxxxxxxxxxxxxxxxxxxxxxx@xxxxxxxxxxxx';

// attendance room id
var ROOM_ID_ALL = 00000000;

// manager room id
var ROOM_ID_MANAGER = 00000000;

// account token
var ACCOUNT_TOKEN_PARUPUNTA = “xxxxxxxxxxxxxxxxxxxxxxxxx”; //パルプン太

// account id
var ACCOUNT_ID_PARUPUNTA = "00000000";

// 0 : from all to all
// 1 : from all to person
// 2 : from peson to person
var EXECUTE_TYPE = 1;

// sheet name
var NAME_ATTENDANCE_MST_SHEET = '打刻対象者マスタ';
var NAME_REST_SHEET = '有休管理';
var NAME_LOG_SHEET = 'ログ';
var PROTECTION_ORNER = '';

// enable google calender flag
var ENABLE_GOOGLE_CALENDER = true;

// enable Gmail flag
var ENABLE_GMAIL = true;

// cell name
var CELL_ACCOUNT_ID = 'B4';
var CELL_USED_REST_NUM = 'D3';
var CELL_REST_INDEX = 'E3';
var CELL_REST = 'E4';
var CELL_ROOM_ID = 'C4';
var CELL_TODAY_DATE = 'A1';
var CELL_TODAY_INDEX = 'A2';
var CELL_CHANGE_DATE = 'A3';
var CELL_CHANGE_INDEX = 'A4';
var CELL_USER_NAME = 'D4';

// embossing master sheet name
var RANGE_EMBOSSING_MST = 'EmbossingMst';

// rest master sheet name
var RANGE_REST_MST = 'RestMst';

// attendance sheet's start cell
var NUM_START_AREA = 5;

// editable day limit
var EDITABLE_LIMIT = 30;

// 入力漏れチェックの対象日(何日前か)
// ※翌日の8時チェックなら対象日を前日(1日前)にする必要あり
var INPUT_MISS_CHECK_DAY = 1;

// message list
var MESSAGE_LIST = {
  'err001' : "エラー発生！\n担当者は対応お願いします！\nエラーメッセージ：",
  'err002’ : "\nパルプンテから入力してね。\n",
}

// parupunta's serif list
var SERIF_LIST = {
  '201' : '[rp aid={1} to={2}-{3}] {4}\nおはよー！{5}で出勤登録したよ。',
  '202' : '[rp aid={1} to={2}-{3}] {4}\nお疲れ！{5}で退勤登録したよ。',
  '203' : '[rp aid={1} to={2}-{3}] {4}\nお休み登録したよ。',
  '204' : '[rp aid={1} to={2}-{3}] {4}\n午前休登録したよ。',
  '205' : '[rp aid={1} to={2}-{3}] {4}\n午後休登録したよ。',
  '211' : '[rp aid={1} to={2}-{3}] {4}\n{5}の出勤時間を{6}に修正したよ。',
  '212' : '[rp aid={1} to={2}-{3}] {4}\n{5}の退勤時間を{6}に修正したよ。',
  '213' : '[rp aid={1} to={2}-{3}] {4}\n{5}にお休み登録したよ。',
  '214' : '[rp aid={1} to={2}-{3}] {4}\n{5}に午前休登録したよ。',
  '215' : '[rp aid={1} to={2}-{3}] {4}\n{5}に午後休登録したよ。',
  '216' : '[rp aid={1} to={2}-{3}] {4}\n{5}',
  '217' : '[rp aid={1} to={2}-{3}] {4}\n{5}に特別休暇登録したよ。',
  '218' : '[rp aid={1} to={2}-{3}] {4}\n{5}に特別休暇(午前)登録したよ。',
  '219' : '[rp aid={1} to={2}-{3}] {4}\n{5}に特別休暇(午後)登録したよ。',
  '220' : '[rp aid={1} to={2}-{3}] {4}\n{5}の休暇を取り消したよ。',
  '221' : '[rp aid={1} to={2}-{3}] {4}\n{5}の代休を{6}に登録したよ。',
  '209' : '有休は残り{1}日だよ。',
  '400' : '[rp aid={1} to={2}-{3}] {4}\nちょっと何言ってるか分からない。',
  '401' : '[rp aid={1} to={2}-{3}] {4}\n5分以上経ってるから再登録出来ないよ！',
  '402' : '[rp aid={1} to={2}-{3}] {4}\n5分以上経ってるから再登録出来ないよ！',
  '403' : '[rp aid={1} to={2}-{3}] {4}\n有休無いから欠勤だよ！',
  '404' : '[rp aid={1} to={2}-{3}] {4}\nもう有休残ってないよ！',
  '405' : '[rp aid={1} to={2}-{3}] {4}\nもう有休残ってないよ！',
  '409' : '有休？んなもんねーよ！',
  '411' : '[rp aid={1} to={2}-{3}] {4}\n{5}の出勤時間は編集可能期限過ぎてるから修正出来ないよ。\nどうしても直したいときは管理部に言ってね。',
  '412' : '[rp aid={1} to={2}-{3}] {4}\n{5}の退勤時間は編集可能期限過ぎてるから修正出来ないよ。\nどうしても直したいときは管理部に言ってね。',
  '413' : '[rp aid={1} to={2}-{3}] {4}\nもう有休残ってないよ！',
  '414' : '[rp aid={1} to={2}-{3}] {4}\nもう有休残ってないよ！',
  '415' : '[rp aid={1} to={2}-{3}] {4}\nもう有休残ってないよ！',
  '420' : '[rp aid={1} to={2}-{3}] {4}\n{5}の休暇は編集可能期限過ぎてるから修正出来ないよ。\nどうしても直したいときは管理部に言ってね。',
  '421' : '[rp aid={1} to={2}-{3}] {4}\n{5}の代休を{6}にとれなかったよ。\nどうしてもとりたいときは管理部に言ってね。',
}

// regular expression pattern
var REGULAR_EXPRESSION = {
  "21" : /^(\#代休\#) \d{4}\/\d{2}\/\d{2} → \d{4}\/\d{2}\/\d{2}/,
  "20" : /^(\#修正\#) \d{4}\/\d{2}\/\d{2} 取消/,
  "18" : /^(\#特別\#) \d{4}\/\d{2}\/\d{2} 午前休/,
  "19" : /^(\#特別\#) \d{4}\/\d{2}\/\d{2} 午後休/,
  "17" : /^(\#特別\#) \d{4}\/\d{2}\/\d{2} 全休/,
  "16" : /(今日|明日|昨日|明後日|一昨日|きのう|あす|あさって|おととい|きょう|([0-1]?[0-9]月[0-3]?[0-9]日)|([0-1]?[0-9]\/[0-3]?[0-9])|(\d{4}\/\d{2}\/\d{2})).*予定/,
  "14" : /^(\#有給\#) \d{4}\/\d{2}\/\d{2} 午前休/,
  "15" : /^(\#有給\#) \d{4}\/\d{2}\/\d{2} 午後休/,
  "13" : /^(\#有給\#) \d{4}\/\d{2}\/\d{2} 全休/,
  "11" : /^(\#修正\#) \d{4}\/\d{2}\/\d{2} ([0-1][0-9]|[2][0-3]):[0-5][0-9]:[0-5][0-9] (出勤|出社)/,
  "12" : /^(\#修正\#) \d{4}\/\d{2}\/\d{2} ([0-1][0-9]|[2][0-3]):[0-5][0-9]:[0-5][0-9] (退勤|退社)/,
  "9" : /(有給|有休).*(教|おしえ)/,
  "5" : /(午後.*(休|やす(ま|み|む)|休暇|無理|ムリ))/,
  "4" : /((午前|朝).*(休|やす(ま|み|む)|休暇|無理|ムリ))|((午後|お昼|(時(頃|ごろ))).*(行|出))/,
  "2" : /(バ[ー〜ァ]*イ|ば[ー〜ぁ]*い|おやすみ|お[つっ]ー|おつ|かえります|さらば|お先|お疲|帰|乙|退勤|ごきげんよ|グ[ッ]?バイ|限界)/,
  "3" : /(全休|(休|やす(ま|み|む))|休暇|自宅作業|無理|ムリ)/,
  "1" : /(モ[ー〜]+ニン|も[ー〜]+にん|おっは|おは|へろ|はろ|ヘロ|morning|ハロ|出[勤社]|(ち.*[ー〜っ]+す)|(業務.*始))/,
}