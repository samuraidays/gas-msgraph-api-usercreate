//メニューを構築する
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('▶OAuth認証')
  .addItem('認証の実行', 'startoauth')
  .addItem('ユーザー新規追加', 'userAdd')
  .addItem('ユーザー情報更新', 'userUpdate')
  .addSeparator()
  .addItem('カレンダー登録', 'createEvent')
  .addSeparator()
  .addItem('ログアウト', 'reset')
  .addToUi();
}
//認証用の各種変数
const sp = PropertiesService.getScriptProperties();
const appid = sp.getProperty('APPID');
const appsecret= sp.getProperty('APPSECRET');
const tokenurl = sp.getProperty('TOKENURL');
const authurl = sp.getProperty('AUTHURL');
const scope = "User.ReadWrite.All offline_access";
const endpoint = "https://graph.microsoft.com";

function startoauth(){
  //UIを取得する
  const ui = SpreadsheetApp.getUi();
  //認証済みかチェックする
  const service = checkOAuth();
  if (!service.hasAccess()) {
    //認証画面を出力
    const output = HtmlService.createHtmlOutputFromFile('template').setHeight(310).setWidth(500).setSandboxMode(HtmlService.SandboxMode.IFRAME);
    ui.showModalDialog(output, 'OAuth2.0認証');
  } else {
    //認証済みなので終了する
    ui.alert("すでに認証済みです。");
  }
}

//アクセストークンURLを含んだHTMLを返す関数
function authpage(){
  const service = checkOAuth();
  const authorizationUrl = service.getAuthorizationUrl();
  const html = "<center><b><a href='" + authorizationUrl + "' target='_blank' onclick='closeMe();'>アクセス承認</a></b></center>"
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
  const service = checkOAuth();
  const isAuthorized = service.handleCallback(request);
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
