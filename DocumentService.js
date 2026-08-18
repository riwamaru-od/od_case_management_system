/**
 * DocumentService.gs
 * 帳票（見積書／請求書／納品書）の実体ファイル操作。
 *  - テンプレートからのコピー生成（初回作成時のみ）
 *  - 再作成時は「別ファイル」ではなく、同一ファイル内でシートを複製する（確定仕様F-18・業務フロー⑤-1）
 *  - フォルダ配置（期／未請求案件・請求中案件）
 *  - PDF出力・印刷用URLの取得
 *
 * 命名規則:
 *   案件ごとに専用フォルダ「{案件番号}_{案件名}」を、書類種別のステージフォルダ配下に作る。
 *   その中に書類ファイルを「{案件番号}_{書類種別}」という名前で1つだけ作成し（初回作成時のみ）、
 *   以降の再作成はこのファイルの中に新しいシートを複製して積み重ねていく。
 *   ファイル内では、現行版のシートを「{書類種別}_最新」、退避された旧版のシートを
 *   「{書類種別}_v{枝番}_{yyyyMMdd}」という名前で保持する（旧版シートは protectSheet_ でロック）。
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

/**
 * テンプレートをコピーして「{案件番号}_{書類種別}」という名前の新規ファイルを作る（初回作成時のみ）。
 * このファイル名（スプレッドシートのタイトル）はファイル自体が存続する限り変わらない
 * （再作成してもファイルは同一のまま、中のシートだけが積み重なっていくため）。
 * 生成直後にファイルオーナーを管理用アカウントへ移譲する（protectSheet_ の実効性を担保するため）。
 */
function createLatestDocument_(docType, caseInfo, stage) {
  const folder = getCaseDocFolder_(docType, caseInfo, stage);
  const templateFile = DriveApp.getFileById(docType.getTemplateFileId());
  const newFile = templateFile.makeCopy(`${caseInfo.caseNo}_${docType.label}`, folder);

  const sheet = SpreadsheetApp.openById(newFile.getId()).getSheets()[0];
  sheet.setName(`${docType.label}${LATEST_SUFFIX}`);

  try {
    transferFileOwnerToAdminAccount_(newFile);
  } catch (e) {
    console.warn(`ファイルオーナーの管理用アカウントへの変更に失敗しました: ${e}`);
  }

  return newFile;
}

/**
 * 再作成処理: ファイルは変えず、ファイル内の現行「最新」シートを複製する。
 *  1. 現行シートをそのまま複製する（データ転記前の状態のまま）
 *  2. 複製元（旧版）シートを枝番付きの名前へリネームし、protectSheet_ でロックする
 *  3. 複製後の新シートを「{書類種別}_最新」として命名・先頭に配置する
 * データ転記（fillQuoteDocument_ 等）は呼び出し元が、この関数が返すファイルに対して
 * 改めて getPrimarySheet_ で「最新」シートを解決して行う。
 *
 * ファイルオーナーは初回作成時（createLatestDocument_）で既に管理用アカウントへ
 * 移譲済みのため、再作成時に改めて移譲する必要はない。
 */
function recreateLatestDocument_(docType, caseInfo) {
  const fileId = extractFileIdFromUrl_(caseInfo[`${docType.key}Link`]);
  const file = DriveApp.getFileById(fileId);
  const ss = SpreadsheetApp.openById(fileId);
  const oldSheet = getPrimarySheet_(file, docType);

  const newSheet = oldSheet.copyTo(ss);

  const nextBranch = countExistingSheetVersions_(ss, docType.label) + 1;
  const dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  oldSheet.setName(`${docType.label}_v${nextBranch}_${dateStr}`);
  try {
    protectSheet_(oldSheet);
  } catch (e) {
    console.warn(`旧シートの保護(protectSheet_)の適用に失敗しました: ${e}`);
  }
  // 注意: ファイル編集権限の降格(restrictFileEditAccessToAdminOnly_)はここでは行わない。
  // このファイルには直後に呼び出し元が新シートへデータ転記を行うため、ここで降格すると
  // 実行者自身がその書き込みに失敗する。降格は呼び出し元（recreateDocumentForCase_）側で
  // データ転記が完了した後に行う。

  newSheet.setName(`${docType.label}${LATEST_SUFFIX}`);
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(1);

  return file;
}

/** ファイル内の「{label}_v数字_日付」シート数をカウントする（次の枝番決定用） */
function countExistingSheetVersions_(ss, label) {
  const prefix = `${label}_v`;
  return ss.getSheets().filter(s => s.getName().indexOf(prefix) === 0).length;
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
 * PDF書き出し用URLを組み立てる共通ヘルパー（PDF_PAGE_RANGES が未定義の書類種別向けの
 * フォールバック。対象シート全体を1ページとして書き出す）。
 * docType を渡さない場合は後方互換として1枚目のシートを対象にする。
 */
function buildPdfExportUrl_(fileId, docType) {
  const sheet = getPrimarySheet_(DriveApp.getFileById(fileId), docType);
  const gid = sheet.getSheetId();
  const params = [
    'format=pdf', `gid=${gid}`, 'size=A4', 'portrait=true', 'fitw=true',
    'gridlines=false', 'printtitle=false', 'sheetnames=false',
    'top_margin=0.5', 'bottom_margin=0.5', 'left_margin=0.5', 'right_margin=0.5',
  ].join('&');
  return `https://docs.google.com/spreadsheets/d/${fileId}/export?${params}`;
}

/** 列名（A, B, ..., Z, AA, ...）を1始まりの列番号に変換する */
function columnLetterToIndex_(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n;
}

/** "A1:H54" のようなA1形式のセル範囲を、0始まり・終端排他のr1/c1/r2/c2に変換する */
function parseA1Range_(a1) {
  const m = String(a1).match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) {
    throw AppError_('INVALID_RANGE', `PDF出力範囲の指定が不正です: ${a1}`);
  }
  return {
    c1: columnLetterToIndex_(m[1]) - 1,
    r1: Number(m[2]) - 1,
    c2: columnLetterToIndex_(m[3]), // 終端は排他的（=1始まりの終端列番号そのもの）
    r2: Number(m[4]),
  };
}

/**
 * 指定したA1形式のセル範囲1つを、1ページ（A4）に収まるPNG画像として書き出す。
 * fitw/fith を両方trueにすることで、この範囲単体が確実に1ページへ収まるようにする。
 */
function exportRangeAsPngBlob_(fileId, docType, a1Range) {
  const sheet = getPrimarySheet_(DriveApp.getFileById(fileId), docType);
  const gid = sheet.getSheetId();
  const box = parseA1Range_(a1Range);
  const params = [
    'format=png', `gid=${gid}`,
    `r1=${box.r1}`, `r2=${box.r2}`, `c1=${box.c1}`, `c2=${box.c2}`,
    'size=A4', 'portrait=true', 'fitw=true', 'fith=true',
    'gridlines=false', 'printtitle=false', 'sheetnames=false',
    'top_margin=0.3', 'bottom_margin=0.3', 'left_margin=0.3', 'right_margin=0.3',
  ].join('&');
  const url = `https://docs.google.com/spreadsheets/d/${fileId}/export?${params}`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw AppError_('PDF_EXPORT_FAILED', buildUserErrorMessage_('PDF出力', 'E102'));
  }
  return response.getBlob();
}

const PDF_PAGE_WIDTH_PT = 595;  // A4 幅（ポイント）
const PDF_PAGE_HEIGHT_PT = 842; // A4 高さ（ポイント）

/**
 * 複数の画像を、Googleスライドを介して1枚ずつのページとして結合し、PDFのBlobとして返す。
 * Apps ScriptにはネイティブなPDF結合機能が無いため、一時的なスライドを作成して
 * 各画像をA4サイズのスライド1枚ずつに敷き詰め、プレゼンテーション全体をPDF書き出し
 * することで実現している（結果、PDF内の文字は画像化され検索・コピー不可になる点に注意）。
 * 一時スライドは書き出し後にゴミ箱へ移動する。
 */
function mergeImagesIntoPdfBlob_(imageBlobs, fileName) {
  const presentation = SlidesApp.create(`_tmp_pdf_merge_${Utilities.getUuid()}`);
  try {
    presentation.setPageSize(PDF_PAGE_WIDTH_PT, PDF_PAGE_HEIGHT_PT);
    const originalFirstSlide = presentation.getSlides()[0];

    imageBlobs.forEach(blob => {
      const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      slide.insertImage(blob, 0, 0, PDF_PAGE_WIDTH_PT, PDF_PAGE_HEIGHT_PT);
    });
    originalFirstSlide.remove();
    presentation.saveAndClose();

    const url = `https://docs.google.com/presentation/d/${presentation.getId()}/export/pdf`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw AppError_('PDF_EXPORT_FAILED', buildUserErrorMessage_('PDF出力', 'E104'));
    }
    return response.getBlob().setName(`${fileName}.pdf`);
  } finally {
    DriveApp.getFileById(presentation.getId()).setTrashed(true);
  }
}

/** PDF_PAGE_RANGES に定義された各ページ範囲を1ページずつ書き出し、1つのPDFへ合成する */
function exportRangesAsMergedPdf_(fileId, docType, fileName) {
  const pageRanges = PDF_PAGE_RANGES[docType.key];
  const imageBlobs = pageRanges.map(a1Range => exportRangeAsPngBlob_(fileId, docType, a1Range));
  return mergeImagesIntoPdfBlob_(imageBlobs, fileName);
}

/** 指定ファイルをPDFとして書き出し、指定フォルダへ保存する */
function exportFileToPdf_(fileId, folder, fileName, docType) {
  const pageRanges = docType && PDF_PAGE_RANGES[docType.key];
  const blob = (pageRanges && pageRanges.length)
    ? exportRangesAsMergedPdf_(fileId, docType, fileName)
    : exportWholeSheetAsPdfBlob_(fileId, docType, fileName);
  return folder.createFile(blob);
}

/** PDF_PAGE_RANGES が無い書類種別向けのフォールバック: シート全体を1ページとして書き出す */
function exportWholeSheetAsPdfBlob_(fileId, docType, fileName) {
  const url = buildPdfExportUrl_(fileId, docType);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw AppError_('PDF_EXPORT_FAILED', buildUserErrorMessage_('PDF出力', 'E101'));
  }
  return response.getBlob().setName(`${fileName}.pdf`);
}
