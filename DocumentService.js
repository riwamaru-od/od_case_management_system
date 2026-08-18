/**
 * DocumentService.gs
 * 帳票（見積書／請求書／納品書）の実体ファイル操作。
 *  - テンプレートからのコピー生成
 *  - 再作成時の旧版リネーム＋新規コピー（確定仕様7章）
 *  - フォルダ配置（期／未請求案件・請求中案件）
 *  - PDF出力・印刷用URLの取得
 *
 * 命名規則:
 *   案件ごとに専用フォルダ「{案件番号}_{案件名}」を、書類種別のステージフォルダ配下に作る。
 *   その中に、現行版を「{書類種別}_最新」、退避された旧版を
 *   「{書類種別}_v{枝番}_{yyyyMMdd}」という名前で保存する。
 */

const LATEST_SUFFIX = '_最新';

function caseFolderName_(caseInfo) {
  return `${caseInfo.caseNo}_${caseInfo.caseName}`;
}

/** 指定書類種別・ステージにおける、この案件専用のフォルダを取得（無ければ作成） */
function getCaseDocFolder_(docType, caseInfo, stage) {
  const periodNumber = getCurrentPeriodNumber_();
  const subfolderName = docType.folderForStage(stage);
  const stageFolder = getDocumentFolder_(docType.folderKind, periodNumber, subfolderName);
  return getOrCreateSubfolder_(stageFolder, caseFolderName_(caseInfo));
}

/** テンプレートをコピーして「{書類種別}_最新」という名前の新規ファイルを作る */
function createLatestDocument_(docType, caseInfo, stage) {
  const folder = getCaseDocFolder_(docType, caseInfo, stage);
  const templateFile = DriveApp.getFileById(docType.getTemplateFileId());
  const newFile = templateFile.makeCopy(`${docType.label}${LATEST_SUFFIX}`, folder);
  return newFile;
}

/**
 * 再作成処理: 既存の「{書類種別}_最新」ファイルを枝番付きでリネームして残し、
 * 新しくテンプレートから複製したファイルを改めて「{書類種別}_最新」と命名する。
 * 戻り値は新しい「最新」ファイル。
 */
function recreateLatestDocument_(docType, caseInfo, stage) {
  const folder = getCaseDocFolder_(docType, caseInfo, stage);
  const latestName = `${docType.label}${LATEST_SUFFIX}`;
  const existingIt = folder.getFilesByName(latestName);

  if (existingIt.hasNext()) {
    const existing = existingIt.next();
    const nextBranch = countExistingVersions_(folder, docType.label) + 1;
    const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
    existing.setName(`${docType.label}_v${nextBranch}_${dateStr}`);
  }

  const templateFile = DriveApp.getFileById(docType.getTemplateFileId());
  return templateFile.makeCopy(latestName, folder);
}

/** フォルダ内の「{label}_v数字_日付」ファイル数をカウントする（次の枝番決定用） */
function countExistingVersions_(folder, label) {
  const prefix = `${label}_v`;
  const it = folder.getFiles();
  let count = 0;
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(prefix) === 0) count++;
  }
  return count;
}

/**
 * 請求書のように、承認後にフォルダを移動する必要がある書類のためのヘルパー。
 * 「{書類種別}_最新」ファイルと、旧版一式（フォルダごと）を新ステージのフォルダへ移動する。
 */
function moveCaseDocFolderToStage_(docType, caseInfo, fromStage, toStage) {
  const fromFolder = getCaseDocFolder_(docType, caseInfo, fromStage);
  const periodNumber = getCurrentPeriodNumber_();
  const toStageFolder = getDocumentFolder_(docType.folderKind, periodNumber, docType.folderForStage(toStage));
  moveFileOrFolder_(fromFolder, toStageFolder);
}

/** Drive上でファイル／フォルダを別フォルダへ移動する（親付け替え方式） */
function moveFileOrFolder_(item, destinationFolder) {
  const parents = item.getParents();
  const isFolder = typeof item.getFiles === 'function';
  if (isFolder) {
    destinationFolder.addFolder(item);
  } else {
    destinationFolder.addFile(item);
  }
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== destinationFolder.getId()) {
      if (isFolder) parent.removeFolder(item); else parent.removeFile(item);
    }
  }
}

/**
 * PDF書き出し用URLを組み立てる共通ヘルパー。
 * 印刷範囲・改ページ位置（見積書=1〜3ページ、請求書/納品書=1〜2ページ）は
 * 各テンプレート側であらかじめ「ファイル > 印刷設定」で設定しておく前提とし、
 * ここでは対象シート（1枚目, gid）とA4・余白などの共通パラメータのみ指定する。
 */
function buildPdfExportUrl_(fileId) {
  const sheet = SpreadsheetApp.openById(fileId).getSheets()[0];
  const gid = sheet.getSheetId();
  const params = [
    'format=pdf', `gid=${gid}`, 'size=A4', 'portrait=true', 'fitw=true',
    'gridlines=false', 'printtitle=false', 'sheetnames=false',
    'top_margin=0.5', 'bottom_margin=0.5', 'left_margin=0.5', 'right_margin=0.5',
  ].join('&');
  return `https://docs.google.com/spreadsheets/d/${fileId}/export?${params}`;
}

/**
 * 印刷用URL（新規タブで開く想定）。
 * 注意: ブラウザのセキュリティ上、別タブ（別オリジン）のウィンドウに対して
 * こちら側のスクリプトから window.print() を強制実行することはできない。
 * そのため実装としては「PDFプレビュー（ブラウザ内蔵ビューア）」を新規タブで開き、
 * ビューア右上の印刷アイコンからそのまま印刷できる状態にする。
 */
function getPrintPreviewUrl_(fileId) {
  return buildPdfExportUrl_(fileId);
}

/** 指定ファイルをPDFとして書き出し、指定フォルダへ保存する */
function exportFileToPdf_(fileId, folder, fileName) {
  const url = buildPdfExportUrl_(fileId);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw AppError_('PDF_EXPORT_FAILED', buildUserErrorMessage_('PDF出力', 'E101'));
  }
  const blob = response.getBlob().setName(`${fileName}.pdf`);
  return folder.createFile(blob);
}
