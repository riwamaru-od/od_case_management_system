/**
 * AdminProxyService.gs
 * サイドバーの書き込み系操作を、常に管理用アカウント権限で実行させるためのプロキシ。
 * WebAppEntry.gs（doPost）とセットで使う。詳細はそちらのコメント参照。
 */

/**
 * WebAppEntry.gs の doPost をサーバー間（UrlFetchApp）で呼び出し、実処理を
 * 管理用アカウント権限で実行させる。
 *
 * 注意: access: ANYONE でデプロイしているため、転送先の doPost 内では
 * Session.getActiveUser() は呼び出し元本人を解決しない（空になる）。
 * そのため本人確認は、呼び出し元自身の（この関数が呼ばれている、本来の）
 * 実行コンテキストで getActiveUserEmail_() を呼んで得たメールアドレスを
 * callerEmail としてリクエストに含め、転送先で ACTIVE_USER_EMAIL_OVERRIDE_
 * として使わせる方式で代替している（Utils.gs 参照）。
 */
function callAsAdmin_(action, params) {
  const url = getProp_(PROP_KEYS.WEBAPP_URL);
  const callerEmail = getActiveUserEmail_();

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ action, params, callerEmail }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  let body;
  try {
    body = JSON.parse(response.getContentText());
  } catch (e) {
    throw AppError_('WEBAPP_BAD_RESPONSE',
      `管理用アカウント経由の処理が想定外の応答を返しました（HTTP ${response.getResponseCode()}）。`
      + `詳細: ${response.getContentText().slice(0, 300)}`);
  }

  if (body.error) {
    const err = new Error(body.error.message);
    err.name = body.error.name || 'AppError';
    throw err;
  }
  return body.result;
}
