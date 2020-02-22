function test() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("アカウント作成");    
  const lastRow1 = sheet.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  Logger.log(lastRow1)
}

function userAdd() {
  //API認証
  const service = checkOAuth();
  
  if (service.hasAccess()) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("アカウント作成");    
    // 作成するユーザの情報を取得する
    var userinfo = getUserInfo(sheet);
    const count = Object.keys(userinfo).length
    // 1つずつユーザ作成を行う
    for(i = 0 ; i < count ; i++ ){
      　// 1人分のユーザ情報を作成する
        const payload = createUserJson(i, userinfo);

        //ユーザ作成APIを実行しユーザを作成する
        const userApiUrl = endpoint + "/beta/users"; 
        const response = graphapicall("POST",userApiUrl,payload);
        const ret = JSON.parse(response.getContentText());
    
        //エラーメッセージを取得し結果セルに記述する
        resultSpredSheet(i, sheet, ret)
        // メール送信する
        sendMail(payload)
    }
    Browser.msgBox("処理が完了しました");
  
  } else {
    Browser.msgBox("認証が実行されていませんよ。");
  }
}

// 作成するユーザの情報を連想配列に入れる
function getUserInfo(sheet) {
  const lastARow = sheet.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const lastBRow = sheet.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  const last1Col = sheet.getRange("A1").getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn(); 
  
  const users = sheet.getRange(lastARow+1, 2, lastBRow-lastARow, last1Col-1).getValues();
  var user={};
  const count=users.length;
  var alldata="";
  for(var i=0;i<count;i++){
    user[i]=users[i];
  }
  return user;
}

// APIを叩いてユーザ作成
function createUserJson(i, userinfo){
  var LastName = userinfo[i][5].toUpperCase();
  var FirstName = userinfo[i][6];
  var username = userinfo[i][7];
  var email = userinfo[i][8];
  var defaultPw = createpwd();
  
  /*
  var depertment = sheet.getRange(row, parseInt(startCol) +  parseInt(7)).getValue();
  if (depertment.length <= 0 ){
  depertment = null;
  }
  */
  
  //拡張属性を取得し、変数へ設定する
  var ext1 = "permanent";
      
  //MS GraphAPI実行
  var userJson =
      {
        accountEnabled: true,
        country: "日本",
        //department: depertment,
        //jobTitle: jobTitle,
        displayName: FirstName + " " + LastName,
        givenName: FirstName,
        mailNickname: username,
        passwordPolicies: "DisablePasswordExpiration",
        "passwordProfile": {
        password: defaultPw,
        forceChangePasswordNextSignIn: false
      },
      surname: LastName,
        //mobilePhone: mobileTel,
        userPrincipalName: email,
          "onPremisesExtensionAttributes":{
            extensionAttribute1: ext1
          },
  };
  return userJson
}

// 結果をスプレッドシートに出力
function resultSpredSheet(i, sheet, ret){  
  const row = i + 3
  if (ret["error"]) {
    sheet.getRange(row,1).setValue(ret.error.message);
  } else {
    sheet.getRange(row,1).setValue("成功");
  }  
}

// メール送信
function sendMail(payload){
  const title = '[Success] CreateUser : ' + payload.userPrincipalName
  const basebody = '## これは自動送信メールです\nユーザが作成されました!'
  const userbody = 'Email: ' + payload.userPrincipalName + '\n' + 'defaultPW: ' + payload.passwordProfile.password
  const lastbody = 'create by ' + Session.getActiveUser();
  const body = basebody + '\n\n' + userbody + '\n\n' + lastbody;
  const toAdr = "hasegawa@kanmu.co.jp";
  const sp = PropertiesService.getScriptProperties();
  const ccAdr = sp.getProperty('ccadr');
  const objArgs = {cc:ccAdr}
  MailApp.sendEmail(toAdr, title, body, objArgs);  
}