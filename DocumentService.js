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
 *
 * ファイルのオーナーは「この処理を実行したアカウント」になる。書類保護を実効あるものに
 * するには管理用アカウントがオーナーである必要があるため、実行アカウントが管理用アカウントで
 * あることを検証する（verifyAdminAccountExecution_ 参照）。
 */
function createLatestDocument_(docType, caseInfo, stage) {
  verifyAdminAccountExecution_(caseInfo.caseNo, `${docType.label}作成`);

  const folder = getCaseDocFolder_(docType, caseInfo, stage);
  const templateFile = getFileByIdSafe_(docType.getTemplateFileId(), `${docType.label}テンプレート`);
  const newFile = templateFile.makeCopy(`${caseInfo.caseNo}_${docType.label}`, folder);

  const sheet = SpreadsheetApp.openById(newFile.getId()).getSheets()[0];
  sheet.setName(`${docType.label}${LATEST_SUFFIX}`);

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
  // 複製直後の新シートは「未承認の新しいドラフト」なので、旧版から引き継がれた
  // 社印を取り除く（再承認時に改めて押印される）
  try {
    removeSealImages_(newSheet, newSheet.getRange(docType.cells().SEAL_IMAGE_RANGE));
  } catch (e) {
    console.warn(`複製シートからの社印除去に失敗しました: ${e}`);
  }
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

/**
 * 期フォルダ配下から、この案件のフォルダを探す。
 * 書類種別によって置き場所が異なる（見積書=未請求案件／請求書=未請求案件・請求中案件／
 * 納品書=期フォルダ直下）ため、期フォルダの直下と各ステージフォルダの両方を探索する。
 * 移動先である「中止案件」フォルダの中は探索対象から除く。
 */
function findCaseDocFolders_(docType, caseInfo, periodNumber) {
  const periodFolder = getOrCreateSubfolder_(getRootFolder_(docType.folderKind), getPeriodFolderName_(periodNumber));
  const targetName = caseFolderName_(caseInfo);
  const found = [];

  const collectFrom = parent => {
    const it = parent.getFoldersByName(targetName);
    while (it.hasNext()) found.push(it.next());
  };

  collectFrom(periodFolder);
  const subfolders = periodFolder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    if (subfolder.getName() === SUBFOLDER.CANCELLED) continue; // 移動先は除外
    collectFrom(subfolder);
  }
  return found;
}

/**
 * 案件中止時に、見積書・請求書・納品書それぞれの案件フォルダを
 * 「{書類種別ルート}/xx期/中止案件/」配下へ移動する。
 * @return {string[]} 移動した内容の説明（操作ログ用）
 */
function moveCaseDocFoldersToCancelled_(caseInfo) {
  const periodNumber = getCurrentPeriodNumber_();
  const movedLabels = [];

  ['quote', 'invoice', 'delivery'].forEach(key => {
    const docType = DOC_TYPES[key];
    try {
      const folders = findCaseDocFolders_(docType, caseInfo, periodNumber);
      if (!folders.length) return;
      const destination = getDocumentFolder_(docType.folderKind, periodNumber, SUBFOLDER.CANCELLED);
      folders.forEach(folder => {
        moveFileOrFolder_(folder, destination);
        movedLabels.push(docType.label);
      });
    } catch (e) {
      // 1種別の移動に失敗しても中止処理自体は止めない（気付けるよう操作ログへ残す）
      console.warn(`${docType.label}フォルダの中止案件フォルダへの移動に失敗しました: ${e}`);
      appendOperationLog_(caseInfo.caseNo, '案件中止（フォルダ移動）',
        `${docType.label}の移動に失敗: ${e && e.message ? e.message : e}`, true);
    }
  });

  return movedLabels;
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
    'horizontal_alignment=CENTER', 'vertical_alignment=TOP',
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
 * セル上に浮いている画像（社印など）が複製先へ引き継がれなかった場合に、
 * 元シートから同じ位置・同じサイズで貼り直す。
 * Sheet.copyTo() は別スプレッドシートへの複製時にセル上の画像を引き継がないことがあり、
 * その場合PDFから社印が欠落してしまうため。
 */
function copyOverGridImagesIfMissing_(sourceSheet, targetSheet) {
  const sourceImages = sourceSheet.getImages();
  if (!sourceImages.length) return;
  if (targetSheet.getImages().length >= sourceImages.length) return; // 既に引き継がれている

  targetSheet.getImages().forEach(img => img.remove()); // 中途半端な引き継ぎを避けて貼り直す
  sourceImages.forEach(img => {
    const anchor = img.getAnchorCell();
    const copied = targetSheet.insertImage(
      img.getBlob(), anchor.getColumn(), anchor.getRow(),
      img.getAnchorCellXOffset(), img.getAnchorCellYOffset()
    );
    copied.setWidth(img.getWidth());
    copied.setHeight(img.getHeight());
  });
}

/**
 * 一時スプレッドシート内に「1ページ分の印刷用シート」を1枚作る。
 * 元シートをまるごと複製したうえで、対象範囲の外側にある行・列を「非表示」にする。
 * （行・列の削除ではなく非表示にするのは、削除すると数式の参照が壊れて #REF! に
 *  なってしまうため。非表示の行・列はPDF出力の対象外になる仕様を利用している。
 *  シートごと複製するため、列幅・行高・結合セル・書式も維持される。
 *  セル上の画像だけは引き継がれないことがあるため copyOverGridImagesIfMissing_ で補う）
 *
 * 社印画像はセル上に浮いている画像であり、非表示行・列の中にあってもPDF出力からは
 * 消えない（非表示はセルの内容にのみ適用され、浮動画像の描画には影響しない）ため、
 * 2ページ目以降にも社印が写り込んでしまう。社印は1ページ目のみに残す。
 */
function buildSinglePagePrintSheet_(tempSs, sourceSheet, a1Range, sheetName, docType, isFirstPage) {
  const copied = sourceSheet.copyTo(tempSs).setName(sheetName);
  copyOverGridImagesIfMissing_(sourceSheet, copied);

  if (!isFirstPage && docType.cells().SEAL_IMAGE_RANGE) {
    removeSealImages_(copied, copied.getRange(docType.cells().SEAL_IMAGE_RANGE));
  }

  // 見積書との差分を示す黄色は社内確認用のため、取引先へ渡すPDFには出さない
  try {
    removeQuoteInvoiceDiffRules_(copied);
  } catch (e) {
    console.warn(`差分ハイライトの除去に失敗しました: ${e}`);
  }

  const box = parseA1Range_(a1Range);
  const maxRows = copied.getMaxRows();
  const maxCols = copied.getMaxColumns();

  // 範囲より後ろ → 前 の順に非表示にする（hideRows/hideColumns は1始まりの絶対位置指定）
  if (maxRows > box.r2) copied.hideRows(box.r2 + 1, maxRows - box.r2);
  if (box.r1 > 0) copied.hideRows(1, box.r1);
  if (maxCols > box.c2) copied.hideColumns(box.c2 + 1, maxCols - box.c2);
  if (box.c1 > 0) copied.hideColumns(1, box.c1);
  return copied;
}

/**
 * PDF_PAGE_RANGES に定義された各ページ範囲を、それぞれ1ページに収めたPDFとして書き出す。
 *
 * 実現方法: Googleスプレッドシートには手動でページ区切りを挿入する機能が無く、
 * PDFエクスポートAPIも1回のリクエストで飛び地の複数範囲を別々のページへ分割できない。
 * 一方で「スプレッドシート全体（gid指定なし）をPDF出力すると、各シートが必ず
 * 新しいページから始まる」という仕様がある。そこで一時スプレッドシートを作り、
 * ページ範囲1つにつき1枚のシートを用意（範囲外は非表示）したうえで、そのファイル
 * 全体をPDF化することで「指定範囲＝1ページ」を実現している。
 * 各シートを1ページへ収めるための縮尺指定は scale=4（ページに合わせる）を使う。
 * ※ fitw は「幅のみ」に合わせる指定で高さが溢れる。fith というパラメータは存在しない。
 * 一時ファイルは書き出し後にゴミ箱へ移動する。
 */
function exportRangesAsPagedPdfBlob_(fileId, docType, fileName) {
  const pageRanges = PDF_PAGE_RANGES[docType.key];
  const sourceSheet = getPrimarySheet_(DriveApp.getFileById(fileId), docType);
  const tempSs = SpreadsheetApp.create(`_tmp_pdf_${Utilities.getUuid()}`);

  try {
    const defaultSheet = tempSs.getSheets()[0];
    pageRanges.forEach((a1Range, i) => {
      buildSinglePagePrintSheet_(tempSs, sourceSheet, a1Range, `page${i + 1}`, docType, i === 0);
    });
    tempSs.deleteSheet(defaultSheet); // 複製したシートだけを残す
    SpreadsheetApp.flush();

    // gid を指定しない = ファイル全体（=全シート）を対象に出力する
    // scale=4 は「ページに合わせる」（幅・高さの両方を1ページに収める）
    // horizontal_alignment=CENTER で、縮小後の内容が左に寄らず左右均等の余白になる
    const params = [
      'format=pdf', 'size=A4', 'portrait=true', 'scale=4',
      'gridlines=false', 'printtitle=false', 'sheetnames=false',
      'pagenum=UNDEFINED', 'attachment=false',
      'horizontal_alignment=CENTER', 'vertical_alignment=TOP',
      'top_margin=0.3', 'bottom_margin=0.3', 'left_margin=0.3', 'right_margin=0.3',
    ].join('&');
    const url = `https://docs.google.com/spreadsheets/d/${tempSs.getId()}/export?${params}`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      throw AppError_('PDF_EXPORT_FAILED', buildUserErrorMessage_('PDF出力', 'E102'));
    }
    return response.getBlob().setName(`${fileName}.pdf`);
  } finally {
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  }
}

/** 指定ファイルをPDFとして書き出し、指定フォルダへ保存する */
function exportFileToPdf_(fileId, folder, fileName, docType) {
  const pageRanges = docType && PDF_PAGE_RANGES[docType.key];
  const blob = (pageRanges && pageRanges.length)
    ? exportRangesAsPagedPdfBlob_(fileId, docType, fileName)
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
