//メニューを構築する
function onOpen(e) {
var ui = SpreadsheetApp.getUi();
ui.createMenu('▶OAuth認証')
.addItem('認証の実行', 'startoauth')
.addItem('ユーザー新規追加', 'userAdd')
.addItem('ユーザー情報更新', 'userUpdate')
.addSeparator()
.addItem('ログアウト', 'reset')
.addToUi();
}
//認証用の各種変数
var sp = PropertiesService.getScriptProperties();
var appid = sp.getProperty('appid');
var appsecret= sp.getProperty('appsecret');
var scope = "User.ReadWrite.All offline_access"
var endpoint = "https://graph.microsoft.com"
var tokenurl = "https://login.microsoftonline.com/knmaad.onmicrosoft.com/oauth2/v2.0/token"
var authurl = "https://login.microsoftonline.com/knmaad.onmicrosoft.com/oauth2/v2.0/authorize"
function startoauth(){
  //UIを取得する
  var ui = SpreadsheetApp.getUi();
  //認証済みかチェックする
  var service = checkOAuth();
  if (!service.hasAccess()) {
    //認証画面を出力
    var output = HtmlService.createHtmlOutputFromFile('template').setHeight(310).setWidth(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    ui.showModalDialog(output, 'OAuth2.0認証');
  } else {
    //認証済みなので終了する
    ui.alert("すでに認証済みです。");
  }
}

//アクセストークンURLを含んだHTMLを返す関数
function authpage(){
  var service = checkOAuth();
  var authorizationUrl = service.getAuthorizationUrl();
  var html = "<center><b><a href='" + authorizationUrl + "' target='_blank' onclick='closeMe();'>アクセス承認</a></b></center>"
  return html;
}
//認証チェック
function checkOAuth() {
  return OAuth2.createService("Microsoft Graph")
  .setAuthorizationBaseUrl(authurl)
  .setTokenUrl(tokenurl)
  .setClientId(appid)
  .setClientSecret(appsecret)
  .setScope(scope)
  .setCallbackFunction("authCallback")　//認証を受けたら受け取る関数を指定する
  .setPropertyStore(PropertiesService.getScriptProperties())  //スクリプトプロパティに保存する
  .setParam("response_type", "code");
}
//認証コールバック
function authCallback(request) {
  var service = checkOAuth();
  Logger.log(request);
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
  return HtmlService.createHtmlOutput("認証に成功しました。ページを閉じてください。");
  } else {
  return HtmlService.createHtmlOutput("認証に失敗しました。");
  }
}

//ログアウト
function reset() {
  checkOAuth().reset();
  SpreadsheetApp.getUi().alert("ログアウトしました。")
}