// Chatwork とスプレッドシートを同期する際、複数のアカウントで連携する方法
//(今回は2つのアカウントを連携する方法を記述)

/* 呼び出し元のアカウントを取得する */
function doGet(e) {
  return makeSessionContent();
}

function makeSessionContent() {
  var result = {'Active User E-mail': Session.getActiveUser().getEmail(),
    'Active User Locale': Session.getActiveUserLocale(),
    'Effective User E-mail': Session.getEffectiveUser().getEmail(),
    'Script Time Zone': Session.getScriptTimeZone(),
    'Temporary Active User Key': Session.getTemporaryActiveUserKey()
  };

  var output = ContentService.createTextOutput(JSON.stringify(result));
  return output.setMimeType(ContentService.MimeType.JSON);
}

function inputCheck() {

  var mySs=SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシートを取得
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var maxRow=mySheet.getRange('F:F').getValues();　 //F列の値を全て取得
  var maxColumn=maxRow.filter(String).length;　　//空白の要素を除いた長さを取得
  var inputFlag=1; //入力が全て完了している=1/完了していない=0
  var email = Session.getActiveUser().getEmail();	//ログイン者のメールアドレスを取得
  Logger.log(email);

  //inputFlagが1のままであれば実行
  if(inputFlag==1 && email == "Googleのメールアドレスを入力"){
    /* スプレッドシートの入力が完了しているかをチェック */
    L:for(var i=3;i<=maxRow;i++){
      for(var j=2;j<=maxColumn;j++){
        if(mySheet.getRange(i,j).getValue()==""){
          inputFlag=0;
          break L;
        }
      }
    }
    /* 送るメッセージを生成 */
    var strBody =
    "[info][title]更新完了[/title]" +
    mySs.getName() + " " +
    mySs.getUrl() + "[/info]"
    
    /* チャットワークにメッセージを送る */
    var cwClient = ChatWorkClient.factory({token: 'APIトークンを入力'}); //チャットワークAPI
    cwClient.sendMessage({
    room_id:ルームIDを入力, //ルームID
    body:strBody
    });
    
   /* 正常に実行できていればメッセージが表示される */
    SpreadsheetApp.getActiveSpreadsheet().toast('処理が完了しました。', '完了', 5);
  }
  
  //inputFlagが1のままであれば実行
  if(inputFlag==1 && email == "Googleのメールアドレスを入力"){
    /* スプレッドシートの入力が完了しているかをチェック */
    L:for(var i=3;i<=maxRow;i++){
      for(var j=2;j<=maxColumn;j++){
        if(mySheet.getRange(i,j).getValue()==""){
          inputFlag=0;
          break L;
        }
      }
    }
    /* 送るメッセージを生成 */
    var strBody =
    "[info][title]更新完了[/title]" +
    mySs.getName() + " " +
    mySs.getUrl() + "[/info]"
    
    /* チャットワークにメッセージを送る */
    var cwClient = ChatWorkClient.factory({token: 'APIトークンを入力'}); //チャットワークAPI
    cwClient.sendMessage({
    room_id:ルームIDを入力, //ルームID
    body:strBody
    });
    
    /* 正常に実行できていればメッセージが表示される */
    SpreadsheetApp.getActiveSpreadsheet().toast('処理が完了しました。', '完了', 5);
  }
}
