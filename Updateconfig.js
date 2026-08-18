/**
 * 過去1ヶ月〜未来6ヶ月の「yyyy/mm 前半」「yyyy/mm 後半」を自動出力・更新する関数
 */
function updateDateRanges() {
  // --- 設定項目 ---
  const SHEET_NAME = 'config'; // 書き込み先のシート名に変更してください
  const START_CELL = 'A1';     // 出力を開始するセル（例: A1から下に展開）
  // ---------------

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log('指定されたシートが見つかりません: ' + SHEET_NAME);
    return;
  }

  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth(); // 0-indexed (0:1月, 11:12月)

  const outputData = [];

  // 過去1ヶ月（-1）から 未来6ヶ月（+6）までの計8ヶ月分をループ
  for (let i = -1; i <= 6; i++) {
    // 対象の年月を計算
    const targetDate = new Date(currentYear, currentMonth + i, 1);
    const year = targetDate.getFullYear();
    // 月を2桁にフォーマット (例: 5 -> "05")
    const month = String(targetDate.getMonth() + 1).padStart(2, '0');

    // 「前半」「後半」の形式で配列に追加
    outputData.push([`${year}/${month} 前半`]);
    outputData.push([`${year}/${month} 後半`]);
  }

  // 出力先セルの範囲を取得して既存内容をクリア後、新しいデータを上書き
  const startRange = sheet.getRange(START_CELL);
  const startRow = startRange.getRow();
  const startCol = startRange.getColumn();

  // 書き込み範囲（行数: outputDataの要素数, 列数: 1列）
  const targetRange = sheet.getRange(startRow, startCol, outputData.length, 1);

  // 古いデータの残存を防ぐため、該当列の下方向をクリア
  sheet.getRange(startRow, startCol, sheet.getLastRow() - startRow + 1, 1).clearContent();

  // 新しいデータを一括書き込み
  targetRange.setValues(outputData);

  Logger.log('日付範囲の更新が完了しました。');
}