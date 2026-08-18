/**
 * ReportService.gs
 * 定期レポート（確定仕様10章）:
 *   前月20日／当月5日 に、対象月の前半・後半の請求案件とステータス一覧を総務へメール送信する。
 *
 * 前提: 「終了予定」「請求予定」列には "YYYY/MM 前半" "YYYY/MM 後半" 形式の値が
 *       プルダウンで入力される想定。実際のプルダウン文言が異なる場合は
 *       billingPeriodLabel_() のフォーマットのみ合わせて修正すればよい。
 */

function billingPeriodLabel_(date, half) {
  const yyyy = date.getFullYear();
  const mm = ('0' + (date.getMonth() + 1)).slice(-2);
  return `${yyyy}/${mm} ${half}`;
}

/** 当日が「前月20日」または「当月5日」であれば、対象月のレポートを送信する */
function sendBillingSummaryReportIfNeeded_() {
  const today = new Date();
  const day = today.getDate();

  let targetMonthDate = null;
  if (day === 20) {
    // 前月20日 → 翌月分のレポート
    targetMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  } else if (day === 5) {
    // 当月5日 → 当月分のレポート
    targetMonthDate = new Date(today.getFullYear(), today.getMonth(), 1);
  } else {
    return;
  }

  const firstHalfLabel = billingPeriodLabel_(targetMonthDate, '前半');
  const secondHalfLabel = billingPeriodLabel_(targetMonthDate, '後半');
  const targets = [firstHalfLabel, secondHalfLabel];

  const cases = getCasesByBillingSchedule_(targets);
  const subject = `【請求予定レポート】${Utilities.formatDate(targetMonthDate, 'Asia/Tokyo', 'yyyy年M月')}分`;
  const body = buildBillingSummaryBody_(targetMonthDate, cases);
  notifyAdminDept_(subject, body);
}

/** 全案件DB（当期分）から、指定した請求予定ラベルに合致する案件を抽出する */
function getCasesByBillingSchedule_(targetLabels) {
  const sheet = getActiveDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CASE_DATA_START_ROW) return [];
  const values = sheet.getRange(CASE_DATA_START_ROW, 1, lastRow - CASE_DATA_START_ROW + 1, CASE_LAST_COL).getValues();

  return values
    .filter(row => targetLabels.indexOf(row[CASE_COLS.BILLING_SCHEDULE - 1]) !== -1)
    .map(row => ({
      caseNo: row[CASE_COLS.CASE_NO - 1],
      clientName: row[CASE_COLS.CLIENT_NAME - 1],
      caseName: row[CASE_COLS.CASE_NAME - 1],
      billingSchedule: row[CASE_COLS.BILLING_SCHEDULE - 1],
      status: row[CASE_COLS.STATUS - 1],
      billingStatus: row[CASE_COLS.BILLING_STATUS - 1],
    }));
}

function buildBillingSummaryBody_(targetMonthDate, cases) {
  const monthLabel = Utilities.formatDate(targetMonthDate, 'Asia/Tokyo', 'yyyy年M月');
  if (cases.length === 0) {
    return `${monthLabel}分の請求予定案件はありません。`;
  }
  const lines = [`${monthLabel}分の請求予定案件一覧（${cases.length}件）`, ''];
  cases.forEach(c => {
    lines.push(`・${c.caseNo}｜${c.clientName}｜${c.caseName}｜請求予定:${c.billingSchedule}｜ステータス:${c.status}｜請求ステータス:${c.billingStatus}`);
  });
  return lines.join('\n');
}
