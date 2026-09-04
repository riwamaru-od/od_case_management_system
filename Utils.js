/**
 * Utils.gs
 * エラーハンドリング、排他制御（LockService）、共通ユーティリティ。
 */

/**
 * アプリ固有のエラー。ユーザー向けメッセージとエラーコードを持つ。
 * サイドバー側では message をそのまま表示する（確定仕様11章のフォーマットに準拠）。
 */
function AppError_(code, message) {
  const err = new Error(message);
  err.name = 'AppError';
  err.code = code;
  return err;
}

/**
 * ユーザー向けの標準エラーメッセージを組み立てる。
 * 例: buildUserErrorMessage_('見積書の作成', 'E001')
 *  → 「見積書の作成に失敗しました。時間を置いて再度やり直すか、管理者にお問い合わせください。[エラーコード：E001]」
 */
function buildUserErrorMessage_(actionLabel, errorCode) {
  return `${actionLabel}に失敗しました。時間を置いて再度やり直すか、管理者にお問い合わせください。[エラーコード：${errorCode}]`;
}

/**
 * LockService で保護しながら fn を実行する共通ラッパー。
 * 失敗時（ロック取得タイムアウト）はユーザー向けエラーを投げる。
 * 例外発生時は呼び出し元で「実行前の状態に戻す」ため、fn 内部の副作用は
 * 「全て検証してから書き込む」設計にすること（部分的な書き込み禁止）。
 */
function withLock_(actionLabel, fn, caseNo) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(10000); // 10秒待って取得できなければタイムアウト
  if (!acquired) {
    throw AppError_('LOCK_TIMEOUT', '他のユーザーが処理中です。しばらくしてから再試行してください。');
  }
  try {
    return fn();
  } catch (e) {
    if (e && e.name === 'AppError') throw e;
    console.error(`[${actionLabel}] Unexpected error: ${e && e.stack ? e.stack : e}`);
    const detail = e && e.message ? ` [詳細: ${e.message}]` : '';
    try {
      appendOperationLog_(caseNo || '', actionLabel, e && e.message ? String(e.message) : String(e), true);
    } catch (logErr) {
      console.error(`操作ログの記録に失敗しました: ${logErr}`);
    }
    throw AppError_('UNEXPECTED', buildUserErrorMessage_(actionLabel, 'E999') + detail);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 「実行者本人」がスクリプトの実行アカウントと異なる場合に、本人のメールアドレスを
 * 一時的に保持する。設定した側が必ず finally で null に戻すこと。
 *
 * 用途は2つ:
 *   1. callAsAdmin_ 経由（Webアプリ・管理用アカウント権限での実行）
 *      → WebAppEntry.gs の doPost が、呼び出し元本人のアドレスを設定する。
 *   2. インストール型 onEdit トリガー（トリガーを登録した管理用アカウントとして実行される）
 *      → CaseService.gs の handleTrackedFieldEdit_ が、実際に編集した本人のアドレスを設定する。
 *
 * 背景: このWebアプリは access: ANYONE でデプロイしている（Workspaceドメインが無いため
 * access: DOMAIN が使えず、DOMAIN指定時は Session.getActiveUser() が呼び出し元の身元を
 * 返す一方、ANYONE指定時はGoogle側がその身元をスクリプトへ渡さない仕様のため）。
 * そのため本人確認は「呼び出し元自身の（本来の）実行コンテキストで getActiveUserEmail_() を
 * 呼んで得た値を、callAsAdmin_ がリクエストに含めて転送する」方式で代替している。
 */
let ACTIVE_USER_EMAIL_OVERRIDE_ = null;

/** 現在の実行ユーザーのメールアドレス */
function getActiveUserEmail_() {
  if (ACTIVE_USER_EMAIL_OVERRIDE_) return ACTIVE_USER_EMAIL_OVERRIDE_;
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw AppError_('NO_USER', 'ユーザー情報を取得できませんでした。Google アカウントでログインしているか確認してください。');
  }
  return email;
}

/** 日時を「yyyy/MM/dd HH:mm」形式の文字列にする */
function formatDateTime_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

/** 3桁ゼロ埋め */
function pad3_(n) {
  return ('000' + n).slice(-3);
}

/**
 * PropertiesService を使ったシンプルな連番カウンター。
 * key ごとに 1 からインクリメントする。呼び出し元で withLock_ に包んで使うこと。
 */
function nextSequence_(key) {
  const props = PropertiesService.getScriptProperties();
  const current = Number(props.getProperty(key) || '0');
  const next = current + 1;
  props.setProperty(key, String(next));
  return next;
}

/** 1始まりの列番号を列名（1→A, 27→AA）に変換する */
function columnIndexToLetter_(index) {
  let letters = '';
  let n = index;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/** シート上で指定した「案件番号」の行番号を探す。見つからなければ null。 */
function findRowByCaseNo_(sheet, caseNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CASE_DATA_START_ROW) return null;
  const values = sheet.getRange(CASE_DATA_START_ROW, CASE_COLS.CASE_NO, lastRow - CASE_DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(caseNo)) {
      return CASE_DATA_START_ROW + i;
    }
  }
  return null;
}
