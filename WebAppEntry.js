/**
 * WebAppEntry.gs
 * サイドバーからの書き込み系操作を、常に管理用アカウント（このWebアプリのデプロイ元アカウント）の
 * 権限で実行するための入り口。
 *
 * 背景: サイドバーの google.script.run 呼び出しは、常に操作した本人の権限で実行される。
 * 案件シートの自動書き込み列（ステータス・書類リンク等）を一般社員の手入力から保護するために
 * 列を保護すると、この「本人権限での実行」が同じ理由でブロックされてしまう
 * （書類ファイルのオーナーが実行者本人になってしまう問題も同じ原因）。
 *
 * そこで appsscript.json の webapp 設定を
 *   { "access": "ANYONE", "executeAs": "USER_DEPLOYING" }
 * としてデプロイし（＝実行ユーザー: 自分（管理用アカウント）、アクセス: ログイン済みの全員）、
 * 書き込み系の各 api_* 関数（SidebarController.gs）からは AdminProxyService.gs の
 * callAsAdmin_() 経由でこの doPost を呼び出す。実処理は常にこちら（管理用アカウント権限）で行う。
 *
 * 注: 本来は access: DOMAIN + 呼び出し元OAuthトークンの転送により、転送先でも
 * Session.getActiveUser() が本人のまま解決される想定だったが、この環境には実際の
 * Google Workspace ドメインが無いため DOMAIN 指定はドメイン所属を解決できずに失敗する。
 * ANYONE 指定に切り替えて解決したが、その代償として Google は doPost 内で
 * Session.getActiveUser() に呼び出し元の身元を渡さなくなる（空文字になる）。
 * そのため本人確認は、callAsAdmin_() が呼び出し元自身のOAuthアクセストークンを
 * callerToken としてリクエストに含め、doPost 側でそのトークンをGoogleの userinfo に
 * 照会して得たメールアドレスを ACTIVE_USER_EMAIL_OVERRIDE_（Utils.gs）にセットし、
 * getActiveUserEmail_() に使わせる方式で代替している。
 * トークンは本人にしか発行されないため、URLを知っている人でも他人にはなりすませない
 * （以前はメールアドレスそのものを受け取っていたため、なりすましが可能だった）。
 *
 * ブラウザ側（JavaScript.html）からは直接叩かない・叩けない設計。呼び出しは常にサーバー間
 * （UrlFetchApp）で行う。
 */

/** doPost が受け付けるアクション名 → 実処理関数のマッピング（params は callAsAdmin_ が渡すオブジェクト） */
const WEBAPP_ACTIONS_ = {
  createQuote: p => createDocumentForCase_('quote', p.caseNo),
  completeQuote: p => completeDocumentForCase_('quote', p.caseNo, p.comment, p.approverEmail),
  approveQuote: p => approveDocumentForCase_('quote', p.caseNo, p.comment),
  rejectQuote: p => rejectDocumentForCase_('quote', p.caseNo, p.comment),
  recreateQuote: p => recreateDocumentForCase_('quote', p.caseNo),
  exportQuotePdf: p => exportDocumentPdfForCase_('quote', p.caseNo),

  createInvoice: p => createDocumentForCase_('invoice', p.caseNo),
  completeInvoice: p => completeDocumentForCase_('invoice', p.caseNo, p.comment),
  approveInvoice: p => approveDocumentForCase_('invoice', p.caseNo, p.comment),
  rejectInvoice: p => rejectDocumentForCase_('invoice', p.caseNo, p.comment),
  recreateInvoice: p => recreateDocumentForCase_('invoice', p.caseNo),
  exportInvoicePdf: p => exportDocumentPdfForCase_('invoice', p.caseNo),

  createDelivery: p => createDeliveryForCase_(p.caseNo),
  exportDeliveryPdf: p => exportDocumentPdfForCase_('delivery', p.caseNo),

  cancelCase: p => cancelCaseForCase_(p.caseNo, p.comment),
  finalApprove: p => finalApproveForCase_(p.caseNo, p.comment),
  restoreCase: p => restoreCaseForCase_(p.caseNo, p.comment),
};

/** Webアプリの唯一の入口。POST本文は {action, params} 形式のJSON。 */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse_({ error: { message: 'リクエストの形式が不正です。', name: 'AppError' } });
  }

  const action = WEBAPP_ACTIONS_[body.action];
  if (!action) {
    return jsonResponse_({ error: { message: `不明な操作です: ${body.action}`, name: 'AppError' } });
  }

  let callerEmail;
  try {
    callerEmail = verifyCallerToken_(body.callerToken);
  } catch (err) {
    return jsonResponse_({ error: { message: err.message, name: err.name || 'AppError' } });
  }

  ACTIVE_USER_EMAIL_OVERRIDE_ = callerEmail;
  try {
    const result = action(body.params || {});
    return jsonResponse_({ result: result });
  } catch (err) {
    return jsonResponse_({
      error: { message: err && err.message ? err.message : String(err), name: (err && err.name) || 'Error' },
    });
  } finally {
    ACTIVE_USER_EMAIL_OVERRIDE_ = null;
  }
}

/** 呼び出し元のアクセストークンの持ち主を確かめる照会先 */
const CALLER_USERINFO_URL_ = 'https://www.googleapis.com/oauth2/v3/userinfo';

/**
 * callAsAdmin_() から渡されたアクセストークンをGoogleに照会し、持ち主のメールアドレスを返す。
 * トークンが無い・無効・期限切れ、またはメールアドレスが確認済みでない場合は例外を投げる。
 * トークンはURLに載せず、Authorizationヘッダーで送る（ログ等に残さないため）。
 */
function verifyCallerToken_(token) {
  if (!token) {
    throw AppError_('UNAUTHENTICATED', '操作した人を確認できませんでした。シートを開き直してから、もう一度お試しください。');
  }
  // 照会のぶん操作が遅くなるため、確認済みのトークンは10分間覚えておく。
  // キーはトークンのハッシュにし、トークンそのものはキャッシュに残さない。
  const cache = CacheService.getScriptCache();
  const cacheKey = 'CALLER_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const response = UrlFetchApp.fetch(CALLER_USERINFO_URL_, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) {
    throw AppError_('UNAUTHENTICATED',
      `操作した人を確認できませんでした（HTTP ${response.getResponseCode()}）。シートを開き直してから、もう一度お試しください。`);
  }
  const info = JSON.parse(response.getContentText());
  if (!info.email || info.email_verified !== true) {
    throw AppError_('UNAUTHENTICATED', '操作した人のメールアドレスを確認できませんでした。管理者へ連絡してください。');
  }
  cache.put(cacheKey, info.email, 600);
  return info.email;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
