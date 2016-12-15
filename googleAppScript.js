var YOUR_API_TOKEN = 'YOUR_SLACK_TOKEN';
var POST_CHANNEL = 'YOUR_SLACK_CHANNEL';
var SHEET_NAME = '休暇チェックリスト';
var CALENDAR_WORD = '[休暇]'; // カレンダーに [休暇] の文字が入っているときだけフック

/*
ライブラリの設定
リソース → ライブラリ → ライブラリを検索に下のIDを入力 → 検索
追加されたライブラリが出てくるので、バージョンを選択する
*/


var _ = underscoreGS; // underscoreGS v3  ID: 1yzaSmLB2aNXtKqIrSZ92SA4D14xPNdZOo3LQRH2Zc6DK6gHRpRK_StrT
SlackApp;             // SlackApp     v22 ID: 1on93YOYfSmV92R5q59NpKmsyWIQD8qnoLYk-gkQBI92C58SPyA2x1-bq


// SlackAPI初回利用時のみ使用 トークン有効期限切れるまでOK
function saveToken() {
  PropertiesService.getScriptProperties().setProperty('token', YOUR_API_TOKEN);
}

function createMessage(aCalData, profile) {
  var title = aCalData.getTitle().replace(CALENDAR_WORD, '');
  var text = profile.superior1 + profile.superior2 + profile.superior3 + '\n'
           + 'お疲れ様です。\n'
           + '本日、'+ title + 'をいただいております。\n'
           + 'メンション付けていただければ、返事はできますので、\n'
           + 'お手数をおかけいたしますがご認識のほどよろしくお願いいたします。';
  return text;
}


function postMessage(aMessage, fromName) {
  var prop = PropertiesService.getScriptProperties().getProperties();

  //slackApp インスタンスの取得
  var slackApp = SlackApp.create(prop.token);

  var channelList = slackApp.channelsList().channels;
  var userList = slackApp.usersList().members;


  var targetChannelName = POST_CHANNEL;
  var targetChannel = null;

  _._find(channelList, function(vals) {
    if (targetChannelName === vals.name) {
      targetChannel = vals;
    }
  });

  _._find(userList, function(vals) {
    if (targetChannelName === vals.name) {
      targetChannel = vals;
    }
  });


  //投稿
  slackApp.postMessage(targetChannel.id, aMessage, {
    username : fromName,
    link_names : 1
  });
}



function main() {
  // スプレットシートの情報を取得
  var profileList = getSS();
  // Logger.log(profileList);

  var profile = {};
  // for(var profileCnt=0; profileCnt<2; profileCnt++) {
  for(var profileCnt=0; profileCnt<profileList.length; profileCnt++) {
    profile = profileList[profileCnt];

    // ターゲットユーザの今日のカレンダー情報を取得
    var todayCal = getCalendarToday(profile.mail);
    // その中で[休暇]の文字が含まれるカレンダータイトルがあれば取得
    var holidayData = checkHolidayEvent(todayCal);

    // データがあれば、休暇文言を作成して、slackにポストする
    var text = '';
    if (holidayData != null) {
      Logger.log('hit');

      text = createMessage(holidayData, profile);
      Logger.log(text);

      if (text != null) {
        Logger.log(text);
        Logger.log(profile.name);
        postMessage(text, profile.name);
      }
    }
  }
}

/* 指定月のカレンダーからイベントを取得する */
function getCalendarToday(_calendarName) {

  var myCal=CalendarApp.getCalendarById(_calendarName); //特定のIDのカレンダーを取得

  var startDate = new Date(); //取得開始日
  var endDate   = new Date();
  startDate.setHours(0);
  startDate.setMinutes(0);
  startDate.setSeconds(0);
  endDate.setHours(23);
  endDate.setMinutes(59);
  endDate.setSeconds(59);

   Logger.log(startDate);
   Logger.log(endDate);



  return myCal.getEvents(startDate,endDate); //カレンダーのイベントを取得

}


function checkHolidayEvent(calList) {

  /* イベントの数だけ繰り返してログ出力 */
  var maxRow = 1;
  var title = '';
  var holidayData = null;
  var HOLIDAY_TEXT = CALENDAR_WORD;
  var holiCnt = 0;
  for each(var evt in calList){

    title = evt.getTitle();

    if (title.indexOf(HOLIDAY_TEXT) != -1) {
      // デバッグ用
      holiCnt++;

      holidayData = evt;
      break;

    }
  }


  return holidayData;
}


// スプレッドシートを取得
function getSS() {
  var mySS=SpreadsheetApp.getActiveSpreadsheet();

  /* シート1とシート2の準備 */
  var sheet1 = mySS.getSheetByName(SHEET_NAME);
  var SSData = sheet1.getDataRange().getValues();
  // Logger.log(SSData);

  var cnt = 0;
  var dataList = [];
  var data = {};
  var record = [];
  var recordName = SSData[1];

  for(var cnt=0; cnt<SSData.length;cnt++) {
    // Logger.log(cnt);

    // ２番めまでは無視
    if (cnt <= 1) { continue; }

    record = SSData[cnt];

    data = {};
    data[recordName[0]] = record[0];
    data[recordName[1]] = record[1];
    data[recordName[2]] = record[2];
    data[recordName[3]] = record[3];
    data[recordName[4]] = record[4];

    dataList.push(data);
  }
  return dataList;
}

