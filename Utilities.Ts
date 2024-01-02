// // With OAuth2 library imported
// // https://github.com/googleworkspace/apps-script-oauth2
// const getAccessTokenForSpreadsheet = function (serviceAccount) {
//     let scope = {
//         spreadsheet: "https://www.googleapis.com/auth/spreadsheets",
//     };
//     var oAuth2Service = OAuth2.createService("ServiceAccount").setTokenUrl(serviceAccount.token_uri).setPrivateKey(serviceAccount.private_key).setIssuer(serviceAccount.client_email).setPropertyStore(PropertiesService.getScriptProperties()).setScope(scope.spreadsheet);
//     if (oAuth2Service.hasAccess()) {
//         return oAuth2Service.getAccessToken();
//     } else {
//         Logger.log("OAuth2 error: " + oAuth2Service.getLastError());
//         return null;
//     }
// };
