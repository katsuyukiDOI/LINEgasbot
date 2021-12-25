// LINE developersのメッセージ送受信設定に記載のアクセストークン
const LINE_TOKEN = 'アクセストークン'; // Messaging API設定の一番下で発行できるLINE Botのアクセストークン（Channel Secretはいらないみたいです。）
const LINE_URL = 'https://api.line.me/v2/bot/message/reply';

//postリクエストを受取ったときに発火する関数
function doPost(e) {

  // 応答用Tokenを取得
  //const replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  // メッセージを取得
  //const userMessage = JSON.parse(e.postData.contents).events[0].message.text;

  var event = JSON.parse(e.postData.contents).events[0];
  // WebHookで受信した応答用Token
  var replyToken = event.replyToken;
  // ユーザーのメッセージを取得
  var userMessage = event.message.text;

   // 応答メッセージの内容
  var messages = [
    {
      type: "text",
      text: "",
    },
  ];
  if (userMessage === "閲覧") {
    recordLineUserId(event.source.userId);   
    recordLineUserIdSheet2A(event.source.userId); 
    incrementC();
　　 messages = confirmation();
    }else if (userMessage === "サンクスレポート") {
    messages = getValueE();
  　}else if (userMessage === "ヒヤリハット") {
    messages = getValueE();
   }else if (userMessage === "ポイント") {
    setNumberF3(event.source.userId);
    messages = getTotalPoint();   
   }else{
  
  //メッセージを改行ごとに分割
  const all_msg = userMessage.split("\n");
  const msg_num = all_msg.length;

  // ***************************
  // スプレットシートからデータを抽出
  // ***************************
  // 1. 今開いている（紐付いている）スプレッドシートを定義
  const sheet     = SpreadsheetApp.getActiveSpreadsheet();
  // 2. ここでは、デフォルトの「シート1」の名前が書かれているシートを呼び出し
  const listSheet = sheet.getSheetByName("シート1");
  // 3. 最終列の列番号を取得
  const numColumn = listSheet.getLastColumn();
  // 4. 最終行の行番号を取得
  const numRow    = listSheet.getLastRow()-1;
  // 5. 範囲を指定（上、左、右、下）
  const topRange  = listSheet.getRange(1, 1, 1, numColumn);      // 一番上のオレンジ色の部分の範囲を指定
  const dataRange = listSheet.getRange(2, 1, numRow, numColumn); // データの部分の範囲を指定
  // 6. 値を取得
  const topData   = topRange.getValues();  // 一番上のオレンジ色の部分の範囲の値を取得
  const data      = dataRange.getValues(); // データの部分の範囲の値を取得
  const dataNum   = data.length +2;        // 新しくデータを入れたいセルの列の番号を取得

  // ***************************
  // スプレッドシートにデータを入力
  // ***************************
  // 最終列の番号まで、順番にスプレッドシートの左からデータを新しく入力
  recordLineUserId(event.source.userId);
  for (let i = 0; i < msg_num; i++) {
    SpreadsheetApp.getActiveSheet().getRange(dataNum, i+1).setValue(all_msg[i]);
  }
  
  const after_msg = {
    'type': 'text',   
    'text': "データを入力しました。",
  }
  messages.push(after_msg);
   }  

  //lineで返答する
  UrlFetchApp.fetch(LINE_URL, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': `Bearer ${LINE_TOKEN}`,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': messages,
    }),
  });
  ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

//シート１のF列にuserIdを登録する関数
function recordLineUserId(userId) {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // F列の空いているセルの行番号を取得する。（F1,F2が既に埋まっていたらnext=3となる）
  var next = activeSheet.getRange("F:F").getValues().filter(String).length + 1;
  Logger.log(next);
  // F列の空いてるセルにユーザーIDを登録する
  activeSheet.getRange(next, 6).setValue(userId);
};

//シート2のA列にuserIdを登録する関数
function recordLineUserIdSheet2A(userId) {
  //シート2を取得
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  // シート２のA列のLINE userIdを取得する
  var userIdList = activeSheet.getRange("A:A").getValues().filter(String).flat();
　// リストになければ登録
　if (!userIdList.includes(userId)) {
  // A列の空いているセルの行番号を取得する。（A1,A2が既に埋まっていたらnext=3となる）
  var next = activeSheet.getRange("A:A").getValues().filter(String).length + 1;
  Logger.log(next);
  // A列の空いてるセルにユーザーIDを登録する
  activeSheet.getRange(next, 1).setValue(userId);
 }
};

//Cセル値（レポート種類）を１にする関数
function incrementC() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // F列の空いているセルの行番号を取得する。（F1,F2が既に埋まっていたらnext=3となる）
  var next = activeSheet.getRange("C:C").getValues().filter(String).length + 1;
  Logger.log(next);
  // C列の空いてるセルに1を登録する
  activeSheet.getRange(next, 3).setValue(1);
}

//IDごとにCセル値（ポイント）を足し算する関数additionNumberC()

//C3セル値を取得する関数
function getNumberC31() {
  //1. 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //2. 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  //3. 指定するセルの範囲（C3）を取得
  var range = sheet.getRange("C3");
  //4. 値を取得する
  var value = range.getValue();
  //ログに出力
  return value;
}

//ボタンを押すと動く関数
function buttonClick() { 
   getHitData();  
}

//検索データ抽出する関数
function getHitData() {
 
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('シート2');
 
  sheet.activate();
 
  let range = sheet.getRange('E2');
  let SearchChar = range.getValue();
 
  sheet.getRange('F2').setValue('=QUERY(シート2!A:B,"SELECT B WHERE A LIKE \'' +SearchChar+'\'")');
 
}

//シート２のBセル値（ID毎）を取得する関数
function getNumberSheet2B(userId) {
//function getNumberC3() {
//function getNumberC3(userId) {
  //シート2を取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];  
   let SearchChar = userId;
  //3. 指定するセルの範囲を取得
  var range = sheet.getRange('=QUERY(A:B,"SELECT B WHERE A LIKE \'' +SearchChar+'\'")');
  
  //var range = sheet.getRange('=QUERY(A:B,"SELECT A,B "WHERE A=' +userId+'")');
  //4. 値を取得する
  var value = range.getValue();
  //ログに出力
  return value;
}

//シート2のF3にID毎のトータルポイントを表示する関数
function setNumberF3(userId) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('シート2'); 
  sheet.activate(); 
   let SearchChar = userId;
    sheet.getRange('F3').setValue('=QUERY(シート2!A:B,"SELECT B WHERE A LIKE \'' +SearchChar+'\'")');
}
//シート2のF3の値を取得する関数
function getNumberF3() {
  //1. 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //2. シート２を取得
  var sheet = spreadsheet.getSheetByName('シート2'); 
  //3. 指定するセルの範囲（C3）を取得
  var range = sheet.getRange("F3");
  //4. 値を取得する
  var value = range.getValue();
  //ログに出力
  return value;
}

//任意のEセル値（概要）を取得する関数getValueE()
function getValueE3() {
  return [    
        
  ];
};

//確認テンプレート
function confirmation(){
  return[
    {
          "type": "template",
          "altText": "晩ご飯をレコメンドします",
          "template": {
            "type": "confirm",
            "text": "閲覧はサンクスレポートですか？ヒヤリハットですか？",
            "actions": [
                {
                  "type": "message", 
                  "label": "サンクスレポート",
                  "text": "サンクスレポート"
                },
                {
                  "type": "message", 
                  "label": "ヒヤリハット",
                  "text": "ヒヤリハット"
                }
            ]
          }
}
]
};

function getValueE() {
  return [
     {  
        "type": "flex",
        "altText": "this is a flex message",
        "contents": {
  "type": "carousel",
  "contents": [
    {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [{
              type: "text",
              text:  String(getNumberE3()),  
              "wrap": true,            
            },]
      }
    },
    {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [{
              type: "text",
              text: String(getNumberE4()), 
              "wrap": true,
        }, ]
      }
    },
     {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [{
              type: "text",
              text: String(getNumberE5()), 
              "wrap": true,
        }, ]
      }
    },
  ]
}
      }
        
  ];
};


//E3セル値を取得する関数
function getNumberE3() {
  //1. 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //2. 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  //3. 指定するセルの範囲（E3）を取得
  var range = sheet.getRange("E3");
  //4. 値を取得する
  var value = range.getValue();
  //ログに出力
  return value;
}

//E4セル値を取得する関数
function getNumberE4() {
  //1. 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //2. 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  //3. 指定するセルの範囲（E4）を取得
  var range = sheet.getRange("E4");
  //4. 値を取得する
  var value = range.getValue();
  //ログに出力
  return value;
}
//E5セル値を取得する関数
function getNumberE5() {
  //1. 現在のスプレッドシートを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //2. 現在のシートを取得
  var sheet = spreadsheet.getActiveSheet();
  //3. 指定するセルの範囲（E5）を取得
  var range = sheet.getRange("E5");
  //4. 値を取得する
  var value = range.getValue();
  //ログに出力
  return value;
}


//ID毎のポイント合計値を表示する関数

//C3セル値のポイントを表示する関数
function getTotalPoint(userId) {
  return [
    {
      type: "flex",
      altText: "スマイルポイント",
      contents: {
        type: "bubble",
        body: {
          type: "box",
          layout: "vertical",
          contents: [
            {
              type: "text",
              text: "獲得スマイルポイント",
              weight: "bold",
              color: "#4169e1",
              size: "sm",
              align: "center",
            },
            {
              type: "text",
              text: String(getNumberF3()),
              //text: String(userId),
              weight: "bold",
              size: "5xl",
              margin: "xxl",
              align: "center",
            },
            {
              type: "separator",
              margin: "xxl",
            },
            {
              type: "box",
              layout: "vertical",
              margin: "xxl",
              spacing: "sm",
              contents: [
                {
              type: "box",
              layout: "horizontal",
              margin: "md",
              contents: [
                {
                  type: "text",
                  text: "ヒアリハット",
                  size: "sm",
                  color: "#4169e1",
                  flex: 0,
                },
                {
                  type: "text",
                  text: "報告 回",
                  color: "#4169e1",
                  size: "sm",
                  align: "end",
                },
              ],
            },
                {
              type: "box",
              layout: "horizontal",
              margin: "md",
              contents: [
                {
                  type: "text",
                  text: "サンクスレポート",
                  size: "sm",
                  color: "#4169e1",
                  flex: 0,
                },
                {
                  type: "text",
                  text: "報告 回",
                  color: "#4169e1",
                  size: "sm",
                  align: "end",
                },
              ],
            },
                {
              type: "box",
              layout: "horizontal",
              margin: "md",
              contents: [
                {
                  type: "text",
                  text: "閲覧回数",
                  size: "sm",
                  color: "#4169e1",
                  flex: 0,
                },
                {
                  type: "text",
                  text: "回",
                  color: "#4169e1",
                  size: "sm",
                  align: "end",
                },
              ],
            },
              ],
            },
            {
              type: "separator",
              margin: "xxl",
            },
            
          ],
        },
        styles: {
          footer: {
            separator: true,
          },
        },
      },
    },
  ];
};
