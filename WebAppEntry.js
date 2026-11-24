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
 * callerToken としてリクエストに含め、doPost 側でそのトークンをGoogleの tokeninfo に
 * 照会し、発行先がこのスクリプトであることと持ち主のメールアドレスを確かめたうえで
 * ACTIVE_USER_EMAIL_OVERRIDE_（Utils.gs）にセットし、
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

/** doPost が受け付けるパラメーターの最大長（異常に大きな入力でシートや通知を汚さないため） */
const WEBAPP_PARAM_MAX_LENGTH_ = {
  caseNo: 50,
  comment: 2000,
  approverEmail: 254,
};

/** Webアプリの唯一の入口。POST本文は {action, params, callerToken} 形式のJSON。 */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e && e.postData ? e.postData.contents : '');
  } catch (err) {
    body = null;
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return jsonResponse_({ error: { message: 'リクエストの形式が不正です。', name: 'AppError' } });
  }

  // プロトタイプ上のプロパティ（toString 等）をアクション名として解決させない
  const action = (typeof body.action === 'string' && Object.prototype.hasOwnProperty.call(WEBAPP_ACTIONS_, body.action))
    ? WEBAPP_ACTIONS_[body.action] : null;
  if (!action) {
    // 送られてきた値はそのまま返さない（応答への任意文字列の反射を避ける）
    return jsonResponse_({ error: { message: '不明な操作です。', name: 'AppError' } });
  }

  let callerEmail;
  let params;
  try {
    callerEmail = verifyCallerToken_(body.callerToken);
    // 社員DBに登録されていない人は、ロール不要の操作（案件中止など）も含めて一切実行させない
    if (!findStaffByEmail_(callerEmail)) {
      throw AppError_('FORBIDDEN', 'このシステムを操作する権限がありません（社員DBに登録されていません）。管理者へ連絡してください。');
    }
    params = normalizeWebAppParams_(body.params);
  } catch (err) {
    return jsonResponse_({ error: { message: err.message, name: err.name || 'AppError' } });
  }

  ACTIVE_USER_EMAIL_OVERRIDE_ = callerEmail;
  try {
    const result = action(params);
    return jsonResponse_({ result: result });
  } catch (err) {
    return jsonResponse_({
      error: { message: err && err.message ? err.message : String(err), name: (err && err.name) || 'Error' },
    });
  } finally {
    ACTIVE_USER_EMAIL_OVERRIDE_ = null;
  }
}

/**
 * doPost に渡されたパラメーターを、各処理が想定する型・長さにそろえる。
 * 想定外の型（オブジェクト・配列など）や長すぎる値は、処理に渡す前に弾く。
 */
function normalizeWebAppParams_(raw) {
  const p = (raw && typeof raw === 'object' && !Array.isArray(raw)) ? raw : {};
  const asText = (key, label) => {
    const v = p[key];
    if (v == null || v === '') return '';
    if (typeof v !== 'string' && typeof v !== 'number') {
      throw AppError_('INVALID_PARAM', `${label}の形式が不正です。`);
    }
    const text = String(v).trim();
    if (text.length > WEBAPP_PARAM_MAX_LENGTH_[key]) {
      throw AppError_('INVALID_PARAM', `${label}が長すぎます（${WEBAPP_PARAM_MAX_LENGTH_[key]}文字以内にしてください）。`);
    }
    return text;
  };
  const caseNo = asText('caseNo', '案件番号');
  if (!caseNo) {
    throw AppError_('INVALID_PARAM', '案件番号が指定されていません。');
  }
  const approverEmail = asText('approverEmail', '承認者のメールアドレス');
  if (approverEmail && !/^[^\s@]+@[^\s@]+$/.test(approverEmail)) {
    throw AppError_('INVALID_PARAM', '承認者のメールアドレスの形式が不正です。');
  }
  return {
    caseNo: caseNo,
    comment: asText('comment', 'コメント'),
    approverEmail: approverEmail,
  };
}

/** アクセストークンの発行先・持ち主を確かめる照会先 */
const CALLER_TOKENINFO_URL_ = 'https://oauth2.googleapis.com/tokeninfo';

/**
 * アクセストークンをGoogleに照会し、その内容（発行先 azp・メールアドレス等）を返す。
 * トークンはURLに載せず、POST本文で送る（アクセスログ等に残さないため）。
 * 無効・期限切れのトークンの場合は null を返す。
 */
function fetchTokenInfo_(token) {
  const response = UrlFetchApp.fetch(CALLER_TOKENINFO_URL_, {
    method: 'post',
    payload: { access_token: token },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) return null;
  try {
    return JSON.parse(response.getContentText());
  } catch (e) {
    return null;
  }
}

/**
 * このスクリプト（Apps Scriptプロジェクト）のOAuthクライアントID。
 * 管理用アカウント自身のトークンを照会して求める。呼び出し元のトークンも同じスクリプトの
 * ScriptApp.getOAuthToken() で発行されるため、正規の呼び出しなら発行先（azp）が一致する。
 */
function getOwnOAuthClientId_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('OWN_OAUTH_CLIENT_ID');
  if (cached) return cached;
  const info = fetchTokenInfo_(ScriptApp.getOAuthToken());
  const clientId = info && (info.azp || info.aud);
  if (!clientId) {
    throw AppError_('UNAUTHENTICATED', '本人確認の準備に失敗しました（このスクリプトのクライアントIDを取得できません）。管理者へ連絡してください。');
  }
  cache.put('OWN_OAUTH_CLIENT_ID', clientId, 21600);
  return clientId;
}

/**
 * callAsAdmin_() から渡されたアクセストークンをGoogleに照会し、持ち主のメールアドレスを返す。
 * 次のいずれかに当てはまる場合は例外を投げる。
 *  - トークンが無い・無効・期限切れ
 *  - トークンの発行先がこのスクリプトではない（他のアプリ向けに発行されたトークンの流用を防ぐ）
 *  - メールアドレスが確認済みでない
 */
function verifyCallerToken_(token) {
  if (!token || typeof token !== 'string' || token.length > 4096) {
    throw AppError_('UNAUTHENTICATED', '操作した人を確認できませんでした。シートを開き直してから、もう一度お試しください。');
  }
  // 照会のぶん操作が遅くなるため、確認済みのトークンは10分間覚えておく。
  // キーはトークンのハッシュにし、トークンそのものはキャッシュに残さない。
  const cache = CacheService.getScriptCache();
  const cacheKey = 'CALLER_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const info = fetchTokenInfo_(token);
  if (!info) {
    throw AppError_('UNAUTHENTICATED', '操作した人を確認できませんでした。シートを開き直してから、もう一度お試しください。');
  }
  const issuedTo = info.azp || info.aud;
  if (!issuedTo || issuedTo !== getOwnOAuthClientId_()) {
    throw AppError_('UNAUTHENTICATED', '操作した人を確認できませんでした（このシステム以外で発行された認証情報です）。');
  }
  const verified = info.email_verified === true || info.email_verified === 'true';
  if (!info.email || !verified) {
    throw AppError_('UNAUTHENTICATED', '操作した人のメールアドレスを確認できませんでした。管理者へ連絡してください。');
  }
  const email = String(info.email).trim().toLowerCase();
  // トークンの残り有効期間より長くは覚えておかない
  const ttl = Math.max(1, Math.min(600, Number(info.expires_in) || 600));
  cache.put(cacheKey, email, ttl);
  return email;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
