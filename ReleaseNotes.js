/**
 * ReleaseNotes.gs
 * システムのバージョン番号と、利用者向けのリリースノートを管理する。
 * メインスプレッドシートの「更新履歴」シートは、ここに書かれた内容から自動で作り直す
 * （シートを直接編集しても、次の更新時に上書きされる。内容はこのファイルで直すこと）。
 *
 * ■ バージョン番号の付け方
 *   ベータ版の間は 0.x.y、正式リリースは 1.0.0 から始める（ベータ版とは番号を分ける）。
 *     - 機能の追加・仕様の変更を含む更新 → 真ん中の数字を上げる（0.9.0 → 0.10.0）
 *     - 不具合の修正・文言の調整だけの更新 → 末尾の数字を上げる（0.9.0 → 0.9.1）
 *   docs/ のマニュアル類の版数（v2.4 など）は文書の版数であり、この番号とは別に管理する。
 *
 * ■ 更新のたびに行うこと
 *   RELEASE_NOTES の先頭に新しい版を1件追加する。追加した版が「現在のバージョン」になり、
 *   スプレッドシートを開いたときに「更新履歴」シートへ自動で反映される。
 *   notes には、利用者から見て何が変わったかを「〜できるようにしました」の形で書く
 *   （内部の実装・CIの変更など、利用者に影響しないものは書かない）。
 *
 *   category: '新機能' | '改善' | '不具合修正' | 'セキュリティ'
 *   audience: '全員' | '管理者'
 */

/** 「更新履歴」シートの名前 */
const RELEASE_NOTES_SHEET_NAME = '更新履歴';

/** シートへ反映済みの内容の指紋（getReleaseNotesSignature_）を記録するスクリプトプロパティのキー */
const RELEASE_NOTES_SYNCED_SIGNATURE_PROP = 'RELEASE_NOTES_SYNCED_SIGNATURE';

/** リリースの段階。正式リリース後は STABLE を使う */
const RELEASE_STAGE = {
  BETA: 'ベータ版',
  STABLE: '正式版',
};

/** 更新履歴（新しい版が先頭）。 */
const RELEASE_NOTES = [
  {
    version: '0.9.0', stage: RELEASE_STAGE.BETA, date: '2026-11-24',
    notes: [
      ['新機能', '全員', '「更新履歴」シートを追加し、システムのバージョンとこれまでの更新内容を確認できるようにしました。'],
      ['改善', '全員', 'シートで案件の行を選んでから、サイドバーに内容が表示されるまでの待ち時間を短くしました（約4秒 → 約2.3秒）。'],
      ['改善', '全員', '書類の承認・PDF出力にかかる時間を短くしました（承認 約14秒 → 約10秒、PDF出力 約15秒 → 約8〜10秒）。'],
      ['不具合修正', '全員', '書類の承認・差し戻し後に、書類ファイルの編集権限を閲覧のみへ切り替える処理が「アクセスが拒否されました」で失敗し、承認後も書類を編集できてしまう不具合を修正しました。書類を再作成すると、再び編集できるようになります。'],
    ],
  },
  {
    version: '0.8.0', stage: RELEASE_STAGE.BETA, date: '2026-11-24',
    notes: [
      ['セキュリティ', '全員', 'サイドバーからの操作で行う本人確認をさらに厳しくし、他のアプリ向けの認証情報を使ったなりすましを防ぐようにしました。'],
      ['セキュリティ', '全員', '社員DBに登録されていない人は、サイドバーから操作できないようにしました。'],
      ['セキュリティ', '全員', 'コメントなどの入力内容が、シートに記録される際に数式として扱われないようにしました。'],
      ['セキュリティ', '管理者', '管理者用メニューの機能は、管理用アカウントとAdminロールの人だけが実行できるようにしました。'],
    ],
  },
  {
    version: '0.7.0', stage: RELEASE_STAGE.BETA, date: '2026-10-15',
    notes: [
      ['新機能', '全員', '通知をメールとChatのどちらで受け取るかを、社員ごとに選べるようにしました（社員DBの「メール通知」「Chat通知」）。'],
      ['新機能', '全員', '書類が差し戻されたとき、その書類を作成した人へ通知が届くようにしました。'],
      ['改善', '全員', '通知メールや定期レポートのデザインを統一しました。'],
      ['新機能', '管理者', '旧「案件番号一覧」から当期の案件を、案件番号を引き継いだまま取り込めるようにしました。'],
      ['改善', '全員', '書類の作成・PDF出力に時間がかかった場合でも、完了後に書類のタブが確実に開くようにしました。'],
      ['不具合修正', '全員', 'サイドバーの自動表示が、設定した本人がスプレッドシートを開いたときだけ動くようにしました。'],
      ['セキュリティ', '全員', 'サイドバーから操作した人を確実に確かめる仕組みに変更し、なりすましを防ぐようにしました。'],
    ],
  },
  {
    version: '0.6.0', stage: RELEASE_STAGE.BETA, date: '2026-09-17',
    notes: [
      ['新機能', '全員', '請求後でも見積書を作り直せるようにしました。作り直した場合、古い請求書・納品書は自動でロックされ、PDF出力もできなくなります。'],
      ['改善', '全員', '取引先名を「社名（担当者名）」のプルダウンから選ぶようにしました。同じ会社で担当者が異なる場合も選び分けられます。'],
      ['改善', '全員', '書類の案件番号・書類番号に「案件No.」「見積書No.」などの見出しを付け、記載位置を変更しました。'],
      ['新機能', '全員', '見積書に、取引先DBの「留意点」を自動で記載するようにしました。'],
      ['改善', '全員', '請求書・納品書の担当社員欄の記載位置を変更しました。'],
    ],
  },
  {
    version: '0.5.0', stage: RELEASE_STAGE.BETA, date: '2026-09-04',
    notes: [
      ['新機能', '全員', '誤って中止・最終承認した案件を、全案件DBで選んでサイドバーの「元に戻す」から進行中へ戻せるようにしました。'],
      ['新機能', '全員', '請求書の内容が見積書と異なるセルを、黄色で表示するようにしました（取引先へ渡すPDFには色は付きません）。'],
      ['新機能', '全員', '見積書の承認依頼で、承認者を指定できるようにしました。承認完了の通知には承認者名を載せます。'],
      ['新機能', '全員', '未承認案件のレポート（毎月25日）、請求予定を過ぎた案件のリマインド（毎月5日）、作成中のまま1週間以上たった書類のお知らせ（毎朝）を追加しました。'],
      ['改善', '全員', '書類の作成時に、入力する欄以外を編集できないようにし、書式や自動記載の欄を誤って書き換えないようにしました。'],
      ['改善', '全員', '見積書の社内作業用エリアは、作成直後から承認後までいつでも編集できるようにしました。'],
      ['改善', '全員', '差し戻し・再作成のあとは、再承認されるまでPDF出力と次の書類の作成をできないようにしました。'],
      ['改善', '全員', 'PDF出力は1ページ（A1:H54）のみとし、社内向けの操作履歴欄などは出力しないようにしました。'],
      ['改善', '全員', 'PDFは最新版だけを残し、古いPDFは自動で削除するようにしました。'],
      ['改善', '全員', '案件を中止したとき、書類一式を「中止案件」フォルダへ自動で移動するようにしました。'],
      ['改善', '全員', '終了予定・請求予定・担当の変更を操作ログに記録し、全案件DBにも反映するようにしました。'],
      ['改善', '全員', '案件番号の並びが崩れた場合、自動で番号順に並べ直すようにしました。'],
      ['改善', '全員', '社印の位置を、テンプレートの行の高さに合わせて自動で調整するようにしました。'],
      ['改善', '管理者', '過去の期のデータは、その期の案件がすべて完了した翌日にアーカイブするようにしました（完了直後の取り消しに対応するため）。'],
      ['不具合修正', '全員', '案件番号が重複することがある不具合を修正しました。'],
      ['不具合修正', '全員', '書類が二重に作られることがある不具合を修正しました。'],
      ['不具合修正', '全員', 'サイドバーのボタンを押しても反応しないことがある不具合を修正しました。'],
      ['不具合修正', '全員', '書類を作り直した際、古い版のシートに社印が残っていた不具合を修正しました。'],
      ['不具合修正', '全員', 'シートを直接編集した操作が、操作ログで「システム(自動実行)」と記録されていた不具合を修正しました。'],
    ],
  },
  {
    version: '0.4.0', stage: RELEASE_STAGE.BETA, date: '2026-08-27',
    notes: [
      ['新機能', '全員', '終了予定・請求予定のプルダウンの選択肢を、月が変わるごとに自動で用意するようにしました。'],
      ['新機能', '全員', '請求予定が今月の案件を、青文字で表示するようにしました。'],
      ['改善', '全員', 'メニューを「ユーザー用システム設定」と「管理者用メニュー」に分けました。'],
      ['改善', '全員', 'サイドバーからの操作を管理用アカウントの権限で実行するようにし、保護されたセルによるエラーを解消しました。'],
      ['改善', '全員', 'フォルダやテンプレートにアクセスできないとき、原因が分かるメッセージを表示するようにしました。'],
      ['不具合修正', '全員', 'PDF出力で、社印が2ページ目以降にも写り込む不具合を修正しました。'],
    ],
  },
  {
    version: '0.3.0', stage: RELEASE_STAGE.BETA, date: '2026-08-18',
    notes: [
      ['新機能', '管理者', '過去の期のデータを別ファイルへ自動でアーカイブし、メインのスプレッドシートが重くならないようにしました。'],
      ['新機能', '管理者', '全案件DBと取引先DBを毎日自動でバックアップするようにしました（30日分を保持）。'],
      ['新機能', '管理者', '管理者用メニューに「案件シートの構成を確認・修復する」を追加しました。'],
      ['改善', '管理者', '毎日の自動処理で1つが失敗しても、残りの処理は続けて実行するようにしました。'],
    ],
  },
  {
    version: '0.2.0', stage: RELEASE_STAGE.BETA, date: '2026-08-18',
    notes: [
      ['新機能', '全員', '案件の中止機能を追加しました（サイドバーの最下部）。'],
      ['新機能', '全員', 'スプレッドシートを開いたときに、サイドバーを自動で表示できるようにしました。'],
      ['新機能', '全員', '最終承認・中止後の案件も、サイドバーで履歴を確認できるようにしました。'],
      ['新機能', '全員', '全操作を記録する「操作ログ」シートを追加しました。'],
      ['改善', '全員', '書類の出力を「PDF出力」に一本化しました。'],
      ['改善', '全員', '承認後の書類と納品書に、社印を自動で押印するようにしました。'],
      ['改善', '全員', '請求書・納品書に、見積書・請求書の通し番号を自動で記載するようにしました。'],
      ['改善', '全員', '書類を作り直すとき、別ファイルではなく同じファイル内に新しいシートを追加するようにしました。'],
      ['改善', '全員', '差し戻された書類や古い版のシートを、自動でロックするようにしました。'],
      ['改善', '全員', '書類のファイル名を「案件番号_書類種別」に統一しました。'],
      ['改善', '全員', '全案件DBで案件の行を選んだ場合も、サイドバーが使えるようにしました。'],
      ['改善', '全員', '総務向けの通知メールを見やすい形式にしました。'],
      ['不具合修正', '全員', '承認フロー・書類管理まわりの不具合を修正しました。'],
    ],
  },
  {
    version: '0.1.0', stage: RELEASE_STAGE.BETA, date: '2026-08-18',
    notes: [
      ['新機能', '全員', 'ベータ版の提供を開始しました。案件の登録と自動採番、見積書・請求書・納品書の作成、承認フローをサイドバーから操作できます。'],
    ],
  },
];

/** 現在のシステムのバージョン（RELEASE_NOTES の先頭の版） */
function getAppVersion_() {
  return RELEASE_NOTES[0].version;
}

/** 表示用のバージョン文字列（例: 「v0.9.0（ベータ版）」） */
function getAppVersionLabel_() {
  const latest = RELEASE_NOTES[0];
  return `v${latest.version}（${latest.stage}）`;
}

/** 「更新履歴」シートの見出し */
const RELEASE_NOTES_HEADERS = ['バージョン', '段階', 'リリース日', '区分', '対象', '内容'];
// サイドバーを開いた状態でも「内容」列が画面に収まる幅にする
const RELEASE_NOTES_COLUMN_WIDTHS = [80, 70, 90, 90, 60, 560];

/** 「更新履歴」シートへ書く行（新しい版が上）を組み立てる */
function buildReleaseNoteRows_() {
  const rows = [];
  RELEASE_NOTES.forEach(release => {
    release.notes.forEach(note => {
      rows.push([`v${release.version}`, release.stage, release.date, note[0], note[1], note[2]]);
    });
  });
  return rows;
}

/**
 * 「更新履歴」シートを RELEASE_NOTES の内容で作り直す（無ければ作成する）。
 * シートの内容は毎回すべて書き直すため、何度実行しても同じ結果になる。
 */
function syncReleaseNotesSheet_() {
  const ss = getMainSpreadsheet_();
  let sheet = ss.getSheetByName(RELEASE_NOTES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RELEASE_NOTES_SHEET_NAME, ss.getSheets().length);
    // 内容はスクリプトから自動で書き直すため、手で編集しようとしたら警告を出す
    sheet.protect()
      .setDescription('更新履歴はシステムが自動で書き換えます。内容は ReleaseNotes.js で編集してください。')
      .setWarningOnly(true);
  }

  const rows = buildReleaseNoteRows_();
  const columnCount = RELEASE_NOTES_HEADERS.length;
  sheet.clear();
  sheet.getRange(1, 1, 1, columnCount)
    .setValues([RELEASE_NOTES_HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8eaed');
  // 内容は文字列としてそのまま表示する（日付への自動変換や数式としての解釈をさせない）
  sheet.getRange(2, 1, rows.length, columnCount)
    .setNumberFormat('@')
    .setValues(rows)
    .setVerticalAlignment('top');
  sheet.getRange(2, columnCount, rows.length, 1).setWrap(true);
  sheet.setFrozenRows(1);
  RELEASE_NOTES_COLUMN_WIDTHS.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  // 版の境目が分かるよう、版ごとに背景色を交互に付ける
  let rowIndex = 2;
  RELEASE_NOTES.forEach((release, i) => {
    sheet.getRange(rowIndex, 1, release.notes.length, columnCount)
      .setBackground(i % 2 === 0 ? '#ffffff' : '#f8f9fa');
    rowIndex += release.notes.length;
  });

  PropertiesService.getScriptProperties()
    .setProperty(RELEASE_NOTES_SYNCED_SIGNATURE_PROP, getReleaseNotesSignature_());
}

/**
 * シートに書く内容と列幅から作る指紋。版番号を変えずに文言を直した場合や
 * 見た目を変えた場合も、指紋が変わるので次に開いたときに書き直される。
 */
function getReleaseNotesSignature_() {
  const source = JSON.stringify([RELEASE_NOTES_HEADERS, RELEASE_NOTES_COLUMN_WIDTHS, buildReleaseNoteRows_()]);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, source, Utilities.Charset.UTF_8);
  return digest.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

/**
 * 「更新履歴」シートが現在の内容を反映していなければ作り直す。
 * スプレッドシートを開いたとき（onOpen）に呼ばれる。同時に開いた人と重ならないよう、
 * ロックが取れなければ何もしない（次に開いた人のときに反映される）。
 */
function syncReleaseNotesSheetIfNeeded_() {
  const synced = PropertiesService.getScriptProperties().getProperty(RELEASE_NOTES_SYNCED_SIGNATURE_PROP);
  const sheetExists = !!getMainSpreadsheet_().getSheetByName(RELEASE_NOTES_SHEET_NAME);
  if (synced === getReleaseNotesSignature_() && sheetExists) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return;
  try {
    syncReleaseNotesSheet_();
  } finally {
    lock.releaseLock();
  }
}

/** メニュー「更新履歴シートを更新する」から呼ばれる */
function syncReleaseNotesSheetManually() {
  const ui = SpreadsheetApp.getUi();
  try {
    withLock_('更新履歴シートの更新', () => syncReleaseNotesSheet_());
  } catch (e) {
    ui.alert(`更新履歴シートの更新に失敗しました。\n\n${e && e.message ? e.message : e}`);
    return;
  }
  ui.alert(`「${RELEASE_NOTES_SHEET_NAME}」シートを更新しました。\n現在のバージョン: ${getAppVersionLabel_()}`);
}
