/**
 * PeriodService.gs
 * 「期」の管理（確定仕様2章）:
 *   - 毎年5/1に期が切り替わる（2026/7時点で17期、2027/5/1から18期）
 *   - 切り替えの1ヶ月前（4/1）に、翌期用のフォルダとシート（表示・全案件DB）を自動作成する
 *   - シートは「原本_表示」「原本_全案件DB」をコピーして「xx期_表示」「xx期_全案件DB」と命名する
 *   - フォルダは 見積書/請求書/納品書 の各ルート配下に「xx期」フォルダを作成する
 */

/** 指定日時点の期番号を返す（例: 2026/7/17 → 17, 2027/5/1 → 18） */
function getPeriodNumber_(date) {
  const y = date.getFullYear();
  const may1ThisYear = new Date(y, PERIOD_SWITCH_MONTH_INDEX, PERIOD_SWITCH_DAY);
  const periodStartYear = date >= may1ThisYear ? y : y - 1;
  return PERIOD_BASE_NUMBER + (periodStartYear - PERIOD_BASE_YEAR);
}

/** 現在の期番号 */
function getCurrentPeriodNumber_() {
  return getPeriodNumber_(new Date());
}

/** 期番号から UI／DBシート名を組み立てる */
function getPeriodSheetNames_(periodNumber) {
  return {
    ui: `${periodNumber}期_表示`,
    db: `${periodNumber}期_全案件DB`,
  };
}

/** 期番号からフォルダ名を組み立てる */
function getPeriodFolderName_(periodNumber) {
  return `${periodNumber}期`;
}

/** 現在アクティブな メイン画面UIシート を取得 */
function getActiveUiSheet_() {
  const names = getPeriodSheetNames_(getCurrentPeriodNumber_());
  const sheet = getMainSpreadsheet_().getSheetByName(names.ui);
  if (!sheet) {
    throw AppError_('SHEET_NOT_FOUND', `シート「${names.ui}」が見つかりません。期の切り替え処理が実行されているか確認してください。`);
  }
  return sheet;
}

/** 現在アクティブな 全案件DBシート を取得 */
function getActiveDbSheet_() {
  const names = getPeriodSheetNames_(getCurrentPeriodNumber_());
  const sheet = getMainSpreadsheet_().getSheetByName(names.db);
  if (!sheet) {
    throw AppError_('SHEET_NOT_FOUND', `シート「${names.db}」が見つかりません。期の切り替え処理が実行されているか確認してください。`);
  }
  return sheet;
}

/**
 * 期または書類種別ごとの通し番号カウンターキーを組み立てる。
 * 例: seqKey_('CASE', 17) → 'SEQ_CASE_17'
 */
function seqKey_(kind, periodNumber) {
  return `SEQ_${kind}_${periodNumber}`;
}

/**
 * 時間主導型トリガーから毎日1回呼び出す想定の関数。
 * 「期の切り替え1ヶ月前(4/1)」に該当する場合、翌期のフォルダ・シートを準備する。
 */
function ensureNextPeriodResourcesIfNeeded() {
  const today = new Date();
  const isPrepDay = today.getMonth() === (PERIOD_SWITCH_MONTH_INDEX - PERIOD_PREP_MONTH_OFFSET) && today.getDate() === PERIOD_SWITCH_DAY;
  if (!isPrepDay) return;

  const nextPeriod = getCurrentPeriodNumber_() + 1;
  withLock_('翌期リソースの準備', () => {
    prepareNextPeriodSheets_(nextPeriod);
    prepareNextPeriodFolders_(nextPeriod);
  });
}

function prepareNextPeriodSheets_(periodNumber) {
  const ss = getMainSpreadsheet_();
  const names = getPeriodSheetNames_(periodNumber);

  if (!ss.getSheetByName(names.ui)) {
    const masterUi = ss.getSheetByName(MASTER_UI_SHEET_NAME);
    const copied = masterUi.copyTo(ss);
    copied.setName(names.ui);
  }
  if (!ss.getSheetByName(names.db)) {
    const masterDb = ss.getSheetByName(MASTER_DB_SHEET_NAME);
    const copied = masterDb.copyTo(ss);
    copied.setName(names.db);
  }
}

function prepareNextPeriodFolders_(periodNumber) {
  const folderName = getPeriodFolderName_(periodNumber);

  const quotePeriodFolder = getOrCreateSubfolder_(getRootFolder_('quote'), folderName);
  getOrCreateSubfolder_(quotePeriodFolder, SUBFOLDER.UNBILLED);

  const invoicePeriodFolder = getOrCreateSubfolder_(getRootFolder_('invoice'), folderName);
  getOrCreateSubfolder_(invoicePeriodFolder, SUBFOLDER.UNBILLED);
  getOrCreateSubfolder_(invoicePeriodFolder, SUBFOLDER.BILLING);

  getOrCreateSubfolder_(getRootFolder_('delivery'), folderName);
}

function getOrCreateSubfolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

/** 書類種別・期・サブフォルダ名から対象フォルダを取得（無ければエラー） */
function getDocumentFolder_(kind, periodNumber, subfolderName) {
  const root = getRootFolder_(kind);
  const periodFolder = getOrCreateSubfolder_(root, getPeriodFolderName_(periodNumber));
  if (!subfolderName) return periodFolder;
  return getOrCreateSubfolder_(periodFolder, subfolderName);
}
