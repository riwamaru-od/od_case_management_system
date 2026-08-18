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
  const msg = buildBillingSummaryMessage_(targetMonthDate, cases);

  withLock_('請求予定レポート送信', () => {
    notifyAdminDept_(msg.subject, msg.body, msg.htmlBody);
    appendOperationLog_('', '請求予定レポート送信', `${cases.length}件`, false);
  });
}

/** 全案件DB（当期分）から、指定した請求予定ラベルに合致する案件を抽出する */
function getCasesByBillingSchedule_(targetLabels) {
  const sheet = getActiveDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CASE_DATA_START_ROW) return [];
  const values = sheet.getRange(CASE_DATA_START_ROW, 1, lastRow - CASE_DATA_START_ROW + 1, CASE_LAST_COL).getValues();

  return values
    .filter(row => targetLabels.indexOf(row[CASE_COLS.BILLING_SCHEDULE - 1]) !== -1
      && row[CASE_COLS.STATUS - 1] !== STATUS.CANCELLED)
    .map(row => ({
      caseNo: row[CASE_COLS.CASE_NO - 1],
      clientName: row[CASE_COLS.CLIENT_NAME - 1],
      caseName: row[CASE_COLS.CASE_NAME - 1],
      billingSchedule: row[CASE_COLS.BILLING_SCHEDULE - 1],
      status: row[CASE_COLS.STATUS - 1],
      billingStatus: row[CASE_COLS.BILLING_STATUS - 1],
    }));
}

/** 請求予定レポートの件名・本文（プレーンテキスト／HTML）を組み立てる */
function buildBillingSummaryMessage_(targetMonthDate, cases) {
  const monthLabel = Utilities.formatDate(targetMonthDate, 'Asia/Tokyo', 'yyyy年M月');
  const subject = `【請求予定レポート】${monthLabel}分`;

  if (cases.length === 0) {
    const body = `${monthLabel}分の請求予定案件はありません。`;
    const htmlBody = buildEmailHtml_(`${monthLabel}分の請求予定レポート`, [{ label: '対象案件', value: 'ありません' }], null, null);
    return { subject, body, htmlBody };
  }

  const lines = [`${monthLabel}分の請求予定案件一覧（${cases.length}件）`, ''];
  cases.forEach(c => {
    lines.push(`・${c.caseNo}｜${c.clientName}｜${c.caseName}｜請求予定:${c.billingSchedule}｜ステータス:${c.status}｜請求ステータス:${c.billingStatus}`);
  });
  const body = lines.join('\n');

  const headerCellStyle = 'padding:6px 10px;text-align:left;font-size:11px;color:#666666;border-bottom:2px solid #dddddd;white-space:nowrap;';
  const cellStyle = 'padding:6px 10px;border-bottom:1px solid #e5e5e5;font-size:12px;color:#222222;';
  const rowsHtml = cases.map(c => `
    <tr>
      <td style="${cellStyle}white-space:nowrap;">${escapeHtmlForMail_(c.caseNo)}</td>
      <td style="${cellStyle}">${escapeHtmlForMail_(c.clientName)}</td>
      <td style="${cellStyle}">${escapeHtmlForMail_(c.caseName)}</td>
      <td style="${cellStyle}white-space:nowrap;">${escapeHtmlForMail_(c.status)}</td>
      <td style="${cellStyle}white-space:nowrap;">${escapeHtmlForMail_(c.billingStatus)}</td>
    </tr>`).join('');
  const htmlBody = `
  <div style="font-family:'Hiragino Kaku Gothic ProN','Meiryo',sans-serif;max-width:640px;margin:0 auto;padding:24px;background-color:#f7f7f7;">
    <div style="background-color:#ffffff;border-radius:8px;padding:24px;border:1px solid #e5e5e5;">
      <h2 style="margin:0 0 16px 0;font-size:16px;color:#222222;">${escapeHtmlForMail_(monthLabel)}分の請求予定案件一覧（${cases.length}件）</h2>
      <table style="width:100%;border-collapse:collapse;">
        <tr>
          <th style="${headerCellStyle}">案件番号</th>
          <th style="${headerCellStyle}">取引先</th>
          <th style="${headerCellStyle}">案件名</th>
          <th style="${headerCellStyle}">ステータス</th>
          <th style="${headerCellStyle}">請求ステータス</th>
        </tr>
        ${rowsHtml}
      </table>
      <p style="margin-top:24px;font-size:11px;color:#999999;">本メールは案件管理・見積書自動作成システムより自動送信されています。</p>
    </div>
  </div>`;
  return { subject, body, htmlBody };
}
