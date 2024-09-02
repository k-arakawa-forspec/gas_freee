const Client_ID = PropertiesService.getScriptProperties().getProperty("clientId");;
const Client_Secret = PropertiesService.getScriptProperties().getProperty("clientSecret");;

function getService() {
  return OAuth2.createService('freee')
  .setAuthorizationBaseUrl('https://accounts.secure.freee.co.jp/public_api/authorize')
  .setTokenUrl('https://accounts.secure.freee.co.jp/public_api/token')
  .setClientId(Client_ID)
  .setClientSecret(Client_Secret)
  .setCallbackFunction('authCallback')
  .setPropertyStore(PropertiesService.getUserProperties())
}

function auth() {
  var driveService = getService();
  if (!driveService.hasAccess()) {
    var authorizationUrl = driveService.getAuthorizationUrl();
    Logger.log('authorizationUrl: ' + authorizationUrl);
    var template = HtmlService.createTemplate(
        '下記のリンクからfreeeを開き、ログインした後に表示されるページの『許可する』ボタンを押下してください。<br><br>' +
        '"認証が成功したのでタブを閉じてください" が表示されたら本ダイアログを閉じてください。<br><br>' +
        '<a href="<?= authorizationUrl ?>" target="_blank">freee</a><br>' +
        '');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showModalDialog(page, 'freeeに移動します');
  } else {
  // ...
  }
}

function logout() {
  var service = getService()
  service.reset();
}

function authCallback(request) {
  var service = getService();
  var isAuthorized = service.handleCallback(request);
  if (isAuthorized){
    return HtmlService.createHtmlOutput('認証が成功したのでタブを閉じてください');
  } else {
    return HtmlService.createHtmlOutput('認証に失敗しています');
  }
}

function getJigyousho() {
  const accessToken = getService().getAccessToken();
  const requestUrl = 'https://api.freee.co.jp/api/1/companies';
  const params = {
    method: 'get',
    headers:{'Authorization':'Bearer ' + accessToken}
  };
  const response = UrlFetchApp.fetch(requestUrl, params);
  Logger.log(response);
  const Sheets = SpreadsheetApp.getActiveSheet();
  Sheets.getRange(1,2).setValue(JSON.parse(response).companies[0].id);
}
