function userAdd() {
  //API認証
  var service = checkOAuth();
  if (service.hasAccess()) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("アカウント作成");    
    
    // データのある最終行を取得する
    const LastRow = getLastLow(sheet);
    
    // 1行ずつセルの情報を取得し、ユーザ作成を行う
    for(i = 0 ; i < LastRow-2 ; i++ ){
      var payload = createUserJson(i, sheet);

      //ユーザ作成APIを実行しユーザを作成する
      var userApiUrl = endpoint + "/beta/users"; 
      var response = graphapicall("POST",userApiUrl,payload);
      var ret = JSON.parse(response.getContentText());
    
 　　  //エラーメッセージを取得し結果セルに記述する
　　　 resultSpredSheet(i, sheet, ret)
      sendMail(payload)
    }
  
    Browser.msgBox("処理が完了しました");
  
  } else {
    Browser.msgBox("認証が実行されていませんよ。");
  }
}

// データのある最終行を取得
function getLastLow(sheet){

  const columnBVals = sheet.getRange("B:B").getValues();
  const LastRow = columnBVals.filter(String).length;
  
  return LastRow
}

// APIを叩いてユーザ作成
function createUserJson(i, sheet){
  var startCol = 1;
  var acheck = sheet.getRange(3,1).getValue();

  if (!acheck){
        
    var row = i + 3
    var jLastName = sheet.getRange(row,parseInt(startCol) + 1).getValue();
    var jFirstName = sheet.getRange(row, parseInt(startCol) + 2).getValue();
    var LastName = sheet.getRange(row,parseInt(startCol) + 6).getValue().toUpperCase();
    var FirstName = sheet.getRange(row, parseInt(startCol) + 7).getValue();
    var username = sheet.getRange(row, parseInt(startCol) + 8).getValue();
    var email = sheet.getRange(row , parseInt(startCol) + 9).getValue();
    //var defaultPw = sheet.getRange(row, parseInt(startCol) + 10).getValue();
    var defaultPw = cpassword();
    
    /*
    var depertment = sheet.getRange(row, parseInt(startCol) +  parseInt(7)).getValue();
    if (depertment.length <= 0 ){
    depertment = null;
    }
    */
    
    /*
    for(extCNT = 1; extCNT <= 15 ; extCNT++){
    eval("var num" + i + "=" + i + ";");
    }
    */
    
    //拡張属性を取得し、変数へ設定する
    var ext1 = "permanent";
    
    /*
    var ext2 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(12))).getValue();
    if (ext2.length <= 0){
    ext2 = null;
    }
    
    var ext3 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(13))).getValue();
    var ext4 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(14))).getValue();
    var ext5 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(15))).getValue();
    var ext6 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(16))).getValue();
    var ext7 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(17))).getValue();
    var ext8 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(18))).getValue();
    var ext9 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(19))).getValue();
    var ext10 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(20))).getValue();
    var ext11 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(21))).getValue();
    var ext12 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(22))).getValue();
    var ext13 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(23))).getValue();
    var ext14 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(24))).getValue();
    var ext15 = sheet.getRange(row, parseInt(startCol) + parseInt(parseInt(25))).getValue();
    */
  }
      
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
  var row = i + 3
  if (ret["error"]) {
    sheet.getRange(row,1).setValue(ret.error.message);
  } else {
    sheet.getRange(row,1).setValue("成功");
  }  
}

// メール送信
function sendMail(payload){
  Logger.log(payload)
  var title = '[Success] CreateUser : ' + payload.userPrincipalName
  var basebody = '## これは自動送信メールです\nユーザが作成されました!'
  var userbody = 'Email: ' + payload.userPrincipalName + '\n' + 'defaultPW: ' + payload.passwordProfile.password
  var lastbody = 'create by ' + Session.getActiveUser();
  var body = basebody + '\n\n' + userbody + '\n\n' + lastbody;
  var toAdr = 'hasegawa@kanmu.co.jp';
  var ccAdr = 'hasegawa@kanmu.co.jp';
  var objArgs = {cc:ccAdr}
  MailApp.sendEmail(toAdr, title, body, objArgs);  
}