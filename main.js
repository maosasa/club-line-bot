//CHANNEL_ACCESS_TOKENを設定
//LINE developerで登録をした、自分のCHANNEL_ACCESS_TOKENを入れて下さい
var CHANNEL_ACCESS_TOKEN = "ACCESS_TOKEN";

// 自分のユーザーIDを指定します。LINE Developersの「Your user ID」の部分です。
var ME = "ID";
var GROUP = "ID"

var line_endpoint = 'https://api.line.me/v2/bot/message/reply';

//カレンダーの情報を取得する
var calendars = CalendarApp.getCalendarById("********@gmail.com");


//全体LINEに送信機能
var group_push_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/*********/edit?usp=sharing")
var group_push_sheet = group_push_ss.getSheetByName("********")
var GP_DATE_ROW = 1
var GP_TIME_ROW = 2
var GP_USER_NAME_ROW = 3
var GP_USER_ID_ROW = 4
var GP_TRIGGER_ID_ROW = 5
var GP_MESSAGE_ROW = 6

//userID管理
var user_id_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/********/edit?usp=sharing")
var user_id_sheet = user_id_ss.getSheetByName("********")
var UI_GRADE_ROW = 1
var UI_LAST_NAME_ROW = 2
var UI_FIRST_NAME_ROW = 3
var UI_ID_ROW = 4

//部練出欠表
var buren_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/********/edit?usp=sharing")
var buren_sheet = buren_ss.getSheetByName("********")
var BUREN_GRADE_ROW = 1
var BUREN_NAME_ROW = 2

//貸切出欠
var kashikiri_ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/********/edit?usp=sharing");
var kashikiri_sheet = kashikiri_ss.getSheetByName("********");
var KASHIKIRI_GRADE_ROW = 1;
var KASHIKIRI_NAME_ROW = 2;

//ポストで送られてくるので、ポストデータ取得
//JSONをパースする
function doPost(e) {
  var json = JSON.parse(e.postData.contents);

  //返信するためのトークン取得
  var reply_token= json.events[0].replyToken;
  if (typeof reply_token === 'undefined') {
    return;
  }

  //送られたLINEメッセージを取得
  var user_message = json.events[0].message.text;
  var groupID = json.events[0].source.groupId;
  var userID = json.events[0].source.userId;
  var type = json.events[0].source.type; //user or group or room

  //返信する内容を作成
  var reply_messages;
  if (type=="user"){//個人トークのみ返信
    if (user_message=='まお') {
      //「まお好き！」って返してくれる
      reply_messages = [ user_message+'好き！',];
    } else if(user_message.match(/こんにちは/)!=null){
      var user_first_name = getUserFirstName(userID)
      if (user_first_name==""){
        reply_messages = ["こんにちは！私はアイスちゃんです！あなたの名前はなんですか？\nユーザー登録#名字\nの形式で教えてくれると嬉しいな！"]
      }else{
        reply_messages = [user_first_name+"さん、こんにちは！"]
      }
    } else if(['ID','id'].indexOf(user_message)!=-1){
      //'ID','id'でIDを返す
      reply_messages = [ userID ];
    }else if(user_message.match(/貸切/)!=null){
      reply_messages = [ "ん？貸切って言った？\n貸切出欠表はこちら！\nhttps://docs.google.com/spreadsheets/d/********/edit#gid=1020227889" ];
    }else if(user_message.match(/ユーザー登録#/)!=null){
      reply_messages = setUserName(user_message,userID);
    }else if(user_message.match(/送信予約#/)!=null){
      reply_messages = setGroupPush(user_message,userID);
    }else if(user_message.match(/送信予約キャンセル/)!=null){
      reply_messages = deleteGroupPush(userID)
    }else if(user_message.match(/送信予約/)!=null){
      reply_messages = ["全体LINEに「送信予約」したい時は、このフォーマットで送ってね！\n送信予約#日付#時間#内容\n例)送信予約#2019-01-01#00:00#あけましておめでとう！"]
    }else{
      //他が入力されたときの処理
      reply_messages = ['ちょっとなに言ってるかわからないです'];
    }
  }
  //グループの場合は'groupID'にのみ反応
  if (type=="group"){
    if (user_message=='groupID'){
      reply_messages = [ groupID ];
    }
    if (user_message.match(/アイスちゃんリマインド/)!=null){
      remindSchedule(true);
    }
  }

  // メッセージを返信
  var messages = reply_messages.map(function (v) {
      return {'type': 'text', 'text': v};
  });
  UrlFetchApp.fetch(line_endpoint, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': messages,
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

//送信するメッセージ定義する関数を作成します。
function createMessage() {
  //メッセージを定義する
  message = "よろしくお願いします！";
  return pushMessage(message, ME);
}

//実際にメッセージを送信する関数を作成します。
function pushMessage(text,to) {
//メッセージを送信(push)する時に必要なurlでこれは、皆同じなので、修正する必要ありません。
//この関数は全て基本コピペで大丈夫です。
  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };

  //toのところにメッセージを送信したいユーザーのIDを指定します。(toは最初の方で自分のIDを指定したので、linebotから自分に送信されることになります。)
  //textの部分は、送信されるメッセージが入ります。createMessageという関数で定義したメッセージがここに入ります。
  var postData = {
    "to" : to,
    "messages" : [
      {
        'type':'text',
        'text':text,
      }
    ]
  };

  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };

  return UrlFetchApp.fetch(url, options);
}


//Momentライブラリ曜日を日本語にセット
Moment.moment.lang('ja', {
    weekdays: ["日曜日","月曜日","火曜日","水曜日","木曜日","金曜日","土曜日"],
    weekdaysShort: ["日","月","火","水","木","金","土"],
});

//次の日の予定リマインド関数
function remindSchedule(is_morning){

    // Date型のオブジェクトの生成
    var date = new Moment.moment();

　　// dateを翌日にセットする
    date.add(1, "days");
　　//date.setDate(date.getDate() + 1);

    //翌日の予定を取得する
    var schedules = calendars.getEventsForDay(date.toDate());

    /*
    * 予定を出力する
    */
    var sendText = new String();
    //sendText += "\n【" + date.format("M月D日(ddd)") + " の予定】";
    //もし予定がない場合は
    if(schedules.length == 0) {
        //sendText += date.format("M月D日(ddd)")+"明日は何も予定がありません";
    } else if(schedules.length > 0 ) {
      // 予定を繰り返し出力する
      for(var i = 0; i < schedules.length; i++) {
        sendText = ""
        var startTime = Moment.moment(schedules[i].getStartTime());
        var endTime = Moment.moment(schedules[i].getEndTime());
        var title = schedules[i].getTitle();
        //var date_D = date.format("M/D");

        if (is_morning){
          if (title.match(/貸切/)!=null){
            //貸切イベント
            var yuushi_text = "";
            if (title.match(/有志/)!=null){ yuushi_text = "有志" }
            sendText += "【" + date.format("M月D日(ddd)") + yuushi_text + "貸切出欠確認】"
            sendText += "\n時間：" + startTime.format("HH:mm") + "〜" + endTime.format("HH:mm");
            sendText += "\n場所：";
            sendText += title.replace("有志","").replace("貸切@","");
            sendText += "\n出席者は";
            sendText += attend_member(startTime.toDate(),"kashikiri");
            sendText += "\nと伺っております。"
            sendText += "\nもし間違いや変更、音源変更がございましたら本日18:00までにご連絡ください。";
          }
        }else{
          if(title.match(/部練/)!=null) {
            sendText +=  "【" + date.format("M月D日(ddd)") + title +  "出席者】"
            sendText += attend_member(date,"buren");
          }else if(title.match(/バッジテスト/)!=null){
            sendText += "【" + date.format("M月D日(ddd)") + title + "】"
            sendText += "\n" + startTime.format("HH:mm") + "〜" + endTime.format("HH:mm")
            //sendText += "\n受験者は以下の通りです"
          }else{
            //その他のイベント
            sendText += "明日の予定です！\n"
            sendText += "" + date.format("M月D日(ddd)")
            if (startTime.isSame(endTime,'minute')){
            }else{
              sendText += startTime.format("HH:mm") + "〜" + endTime.format("HH:mm");
            }
            sendText += "\n"+title
          }

        }
        if(sendText!=""){
            pushMessage(sendText,GROUP);//部のLINEに投稿
        }
      }
    }

  return 0;
}

function attend_member(date,type){
  var attendance_row = 0;
  var sendText = new String();
  if (type == "kashikiri"){
    sheet = kashikiri_sheet
    GRADE_ROW = KASHIKIRI_GRADE_ROW
    NAME_ROW = KASHIKIRI_NAME_ROW
    is_same_key = 'minute'
  }else if(type == "buren"){
    sheet = buren_sheet
    GRADE_ROW = BUREN_GRADE_ROW
    NAME_ROW = BUREN_NAME_ROW
    is_same_key = 'day'
  }

  var days_on_sheet = sheet.getRange(1,1,1,50).getValues()
  for (var i = 2;i<100;i++){
    if(Moment.moment(sheet.getRange(1,i).getValue()).isSame(date,is_same_key)){
      attendance_row = i
      break;
    }
  }
  if (attendance_row == 0){
    sendText += "\n......未登録......";
    return sendText;
  }

  var grades = sheet.getRange(3,GRADE_ROW,50).getValues();
  var names = sheet.getRange(3,NAME_ROW,50).getValues();
  var attendance = sheet.getRange(3,attendance_row,50).getValues();
  var first = true;
  var tmp_grade = 0.0;
  var undecidedMembers ="";
  var undecidedFirst = true;
  for (i = 0; i<50;i++){
    if (names[i] != ""){
    if (attendance[i]=="出席" || attendance[i].toString().match(/早退/)!=null ||attendance[i].toString().match(/途中参加/)!=null){//　出席の人
      if (parseInt(grades[i])!=parseInt(tmp_grade)){
        first = true; //学年が切り替わった時改行
        tmp_grade = grades[i];
      }
      if (first){
        sendText += "\n"; //始めの人は改行
        first = false;
      }else{
        sendText += ","; //始め以外はコンマ
      };
      sendText += names[i];//名前を追加
      if(attendance[i].toString().match(/早退/)!=null ||attendance[i].toString().match(/途中参加/)!=null){
        sendText += "("+attendance[i]+")";//(遅刻or途中参加)を追加
      }
    }else if (names[i]=='他大生'){
      if (attendance[i]!=""){
      sendText += "\n[他大]";
      sendText +=attendance[i];
      }
    }else if (attendance[i]=="未定"||attendance[i]==""){
      if (undecidedFirst){
        undecidedMembers += "\n[未定]"; //始めの人は[未定]
        undecidedFirst = false;
      }else{
        undecidedMembers += ","; //始め以外はコンマ
      };
      undecidedMembers += names[i]
    }
    }
  }
  sendText += undecidedMembers;
  return sendText;
}


function morningRemind(){
  remindSchedule(true)
  return 0;
}

function eveningRemind(){
  remindSchedule(false)
  return 0;
}


function setGroupPush(user_message,userID){
  var user_name = getUserLastName(userID)
  if(user_name == ""){
    return ["ユーザー登録が済んでいない方は送信予約できません。ごめんなさい><"]
  }
  var group_push_command=user_message.split("#")
  var group_push_day = ""+group_push_command[1]
  var group_push_times = group_push_command[2].split(":")
  var group_push_hour = Number(group_push_times[0])
  var group_push_min = Number(group_push_times[1])
  var group_push_message = group_push_command[3]
  group_push_sheet.insertRowAfter(1);
  group_push_sheet.getRange(2,GP_DATE_ROW).setValue(group_push_day);
  group_push_sheet.getRange(2,GP_TIME_ROW).setValue(group_push_command[2]);
  group_push_sheet.getRange(2,GP_USER_NAME_ROW).setValue(user_name);
  group_push_sheet.getRange(2,GP_USER_ID_ROW).setValue(userID);
  group_push_sheet.getRange(2,GP_MESSAGE_ROW).setValue(group_push_message);
  pushMessage(""+group_push_day+" "+group_push_command[2]+"\n"+group_push_message+"\nで予約しました！",userID);
  setTrigger(group_push_day,group_push_hour,group_push_min)
  return []
}

function deleteGroupPush(userID){
  var reply_message = [""]
  var delete_row_list = []
  for (var i = 1;i<50;i++){
    if(group_push_sheet.getRange(i,GP_USER_ID_ROW).getValue()==userID){
      day = new Moment.moment(group_push_sheet.getRange(i,GP_DATE_ROW).getValue()).format("YYYY/MM/DD");//フォーマット変換うまくいかない
      time = new Moment.moment(group_push_sheet.getRange(i,GP_TIME_ROW).getValue()).format("hh:mm");
      message = group_push_sheet.getRange(i,GP_MESSAGE_ROW).getValue();
      trigger_id = group_push_sheet.getRange(i,GP_TRIGGER_ID_ROW).getValue();
      reply_message.push('送信予約#'+day+'#'+time+'#'+message)
      Logger.log(reply_message)
      deleteTrigger(trigger_id)
      delete_row_list.push(i)
    }
  }
  var cnt = 0
  for (i in delete_row_list){
    group_push_sheet.deleteRows(delete_row_list[i-cnt])
    cnt+=1
  }
  Logger.log(delete_row_list)
  if(reply_message.length == 1){
    reply_message = ['このアカウントから送信予約されているメッセージはありませんでした！']
  }else{
    reply_message.shift()
    reply_message.push('送信予約を全て取り消しました！復元したい場合は上のメッセージをそのままコピペして送ってね')
  }
  return reply_message
}
  

function setTrigger(date,h,m) {
 var triggerDay = new Date(date); 
 triggerDay.setHours(h);
 triggerDay.setMinutes(m);
 var trigger =  ScriptApp.newTrigger("groupPush").timeBased().at(triggerDay).create();
 group_push_sheet.getRange(2,GP_TRIGGER_ID_ROW).setValue(trigger.getUniqueId());
}

function groupPush(e){
  for (var i = 1;i<50;i++){
    if(group_push_sheet.getRange(i,GP_TRIGGER_ID_ROW).getValue()==e.triggerUid){
      var userDataRow = i
      break;
    }
  }
 var message = group_push_sheet.getRange(userDataRow,GP_MESSAGE_ROW).getValue();
 var user_name = group_push_sheet.getRange(userDataRow,GP_USER_NAME_ROW).getValue();
  message += "\n("+user_name+"より送信)"
 group_push_sheet.deleteRows(userDataRow)
 deleteTrigger(e.triggerUid)
 return pushMessage(message,UTFS2019);
}

function deleteTrigger(uniqueId) {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() == uniqueId) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function getUserLastName(userID){
  var user_id_data = user_id_sheet.getRange(2,2,50,3).getValues();
  var user_name = ""
  const USER_ID_ROW = UI_ID_ROW -2
  const USER_NAME_ROW = UI_LAST_NAME_ROW -2
  for(var i in user_id_data){
    if(user_id_data[i][USER_ID_ROW]==userID){
      user_name = user_id_data[i][USER_NAME_ROW]
    }
  }
  return user_name
}
function getUserFirstName(userID){
  var user_id_data = user_id_sheet.getRange(2,2,50,3).getValues();
  var user_name = ""
  const USER_ID_ROW = UI_ID_ROW -2
  const USER_NAME_ROW = UI_FIRST_NAME_ROW -2
  for(var i in user_id_data){
    if(user_id_data[i][USER_ID_ROW]==userID){
      user_name = user_id_data[i][USER_NAME_ROW]
    }
  }
  return user_name
}

function setUserName(user_message,userID){
  var user_name = user_message.split("#")[1]
  //スプレッドシートにユーザID登録
  for (var i = 1;i<50;i++){
    if(user_id_sheet.getRange(i,UI_LAST_NAME_ROW).getValue()==user_name){
      var registered_id = user_id_sheet.getRange(i,UI_ID_ROW).getValue()
      if(registered_id==""){
        user_id_sheet.getRange(i,UI_ID_ROW).setValue(userID)
        return [user_name + "さんとして登録しました！"]
      }else if(registered_id==userID) {
        return [user_name + "さんは登録済です"]
      }else{
        return [user_name + "さんは別のIDで登録済です",
               "変更したい場合は管理者に問い合わせて下さい"]
      }
    }
  }
  return [user_name + "さんを名簿から見つけることができませんでした"]
}
