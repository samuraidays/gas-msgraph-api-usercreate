//APIを叩くルーチン
function graphapicall(method, eUrl, sdata) {
  //Graph APIサービスを取得する
  const service = checkOAuth();
  
  if (service.hasAccess()) {
    //HTTP通信
    const response = UrlFetchApp.fetch(eUrl, {
      headers: {
        Authorization: "Bearer " + service.getAccessToken()
      },
      method: method,
      contentType: "application/json",
      payload:JSON.stringify(sdata),
      muteHttpExceptions: true,
    });
    
    return response;
  }else{
    //エラーを返す（認証が実行されていない場合）
    return "error";
  }
}
