//APIを叩くルーチン
function graphapicall(method, eUrl, sdata) {
  //Graph APIサービスを取得する
  var service = checkOAuth();
  
  if (service.hasAccess()) {
    /*
    var payloadData = {
      values:sdata
    }
    
    */
    
    //HTTP通信
    var response = UrlFetchApp.fetch(eUrl, {
      headers: {
        Authorization: "Bearer " + service.getAccessToken()
      },
      method: method,
      contentType: "application/json",
      payload:JSON.stringify(sdata),
      muteHttpExceptions: true,
    });
    
    //Browser.msgBox(JSON.stringify(sdata));
    //取得した値を返す
    //Browser.msgBox(response.getResponseCode());
    //var ret = JSON.parse(response.getContentText());
    
    return response;
  }else{
    //エラーを返す（認証が実行されていない場合）
    return "error";
  }
}
