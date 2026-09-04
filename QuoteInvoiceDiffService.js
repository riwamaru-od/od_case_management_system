/**
 * QuoteInvoiceDiffService.gs
 * 見積書と請求書の内容差分を、請求書側のセル背景色（黄色）でリアルタイムに可視化する。
 *
 * 仕組み:
 *   1. 請求書ファイル内に非表示の比較用シート（QUOTE_INVOICE_DIFF_SHEET_NAME）を作り、
 *      請求書の作成・再作成時点における見積書の値を「同一番地」へ書き込んでおく。
 *   2. 請求書本体（最新シート）に条件付き書式を設定し、比較用シートと値が異なるセルを
 *      黄色にする。条件付き書式はスプレッドシートの機能なので、利用者が請求書を編集した
 *      瞬間に色が付き、見積書と同じ内容に直せば即座に色が消える。
 *
 * この方式を採るのは、請求書が案件ごとに別ファイルであり、ファイルごとにスクリプトの
 * onEdit トリガーを設置する方式では1ユーザーあたりのトリガー数上限に抵触するため。
 */

/**
 * 請求書ファイルに対して、見積書との差分ハイライトを設定（または再設定）する。
 * 請求書の作成時・再作成時（fillInvoiceDocument_ の最後）に呼ぶ。
 * @param {Sheet} invoiceSheet 請求書の「最新」シート
 * @param {string} quoteFileId 比較元となる見積書ファイルのID
 */
function applyQuoteInvoiceDiffHighlight_(invoiceSheet, quoteFileId) {
  const quoteSheet = getPrimarySheet_(DriveApp.getFileById(quoteFileId), DOC_TYPES.quote);
  const diffSheet = getOrCreateDiffSheet_(invoiceSheet.getParent());

  // 見積書の現在値を、比較用シートの同一番地へ複製する
  QUOTE_INVOICE_DIFF_RANGES.forEach(a1 => {
    diffSheet.getRange(a1).setValues(quoteSheet.getRange(a1).getValues());
  });

  applyDiffConditionalFormatRules_(invoiceSheet, diffSheet);
}

/**
 * 比較用シートを取得する（無ければ作成して非表示にする）。
 * シート保護はあえて掛けていない。中身は請求書の作成・再作成のたびに書き直されるため、
 * 仮に編集されても影響は「ハイライトが一時的に正しくなくなる」だけで、
 * 保護APIの失敗（この環境では過去に頻発）を持ち込むリスクの方が大きいため。
 */
function getOrCreateDiffSheet_(ss) {
  let sheet = ss.getSheetByName(QUOTE_INVOICE_DIFF_SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.insertSheet(QUOTE_INVOICE_DIFF_SHEET_NAME);
  sheet.hideSheet();
  return sheet;
}

/**
 * 請求書シートへ「比較用シートと値が違えば黄色」の条件付き書式を設定する。
 * 既存ルールのうち本機能が設定したものだけを差し替え、
 * 利用者が独自に設定した他の条件付き書式は残す。
 *
 * 重要: Googleスプレッドシートの条件付き書式のカスタム数式は、他シートを
 * 「'シート名'!A1」の形で直接参照できない（ルールは作成できるが、常に不成立となり
 * 色が付かない）。他シートを参照するには INDIRECT() で文字列から組み立てる必要がある。
 * そのため各セルの番地を ADDRESS(ROW(), COLUMN()) で組み立て、自セルと比較用シートの
 * 同一番地を INDIRECT で読んで突き合わせている。
 * この数式は範囲によらず同一のため、全範囲をまとめて1つのルールとして設定する。
 */
function applyDiffConditionalFormatRules_(invoiceSheet, diffSheet) {
  const ranges = QUOTE_INVOICE_DIFF_RANGES.map(a1 => invoiceSheet.getRange(a1));
  const formula = buildDiffConditionalFormula_(diffSheet.getName());

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(QUOTE_INVOICE_DIFF_COLOR)
    .setRanges(ranges)
    .build();

  const kept = keepNonDiffConditionalFormatRules_(invoiceSheet);
  invoiceSheet.setConditionalFormatRules(kept.concat([rule]));
}

/** 差分ハイライトのカスタム数式を組み立てる */
function buildDiffConditionalFormula_(diffSheetName) {
  const address = 'ADDRESS(ROW(),COLUMN())';
  return `=INDIRECT(${address})<>INDIRECT("'${diffSheetName}'!"&${address})`;
}

/**
 * シートの条件付き書式のうち、本機能が設定した差分ハイライトを除いたルールを返す。
 * 利用者がテンプレート側で独自に設定した条件付き書式は残す。
 */
function keepNonDiffConditionalFormatRules_(sheet) {
  return sheet.getConditionalFormatRules().filter(rule => !isQuoteInvoiceDiffRule_(rule));
}

/**
 * 条件付き書式のルールが、本機能が設定した差分ハイライトかどうかを判定する。
 * 数式の中で比較用シートを参照しているかどうかで見分ける
 * （対象範囲での判定だと、利用者が同じ範囲に設定した独自ルールまで消してしまうため）。
 */
function isQuoteInvoiceDiffRule_(rule) {
  const condition = rule.getBooleanCondition();
  if (!condition) return false;
  if (condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return false;
  const values = condition.getCriteriaValues();
  return values.length > 0 && String(values[0]).indexOf(QUOTE_INVOICE_DIFF_SHEET_NAME) !== -1;
}

/**
 * 差分ハイライトの条件付き書式を取り除く。
 * PDF出力用の一時シートに対して使う。差分の黄色は社内確認用であり、
 * 取引先へ渡すPDFに出てはいけないため（加えて、一時スプレッドシートには
 * 比較用シートが存在せず、参照が壊れた状態のルールを残すことにもなるため）。
 */
function removeQuoteInvoiceDiffRules_(sheet) {
  sheet.setConditionalFormatRules(keepNonDiffConditionalFormatRules_(sheet));
}
