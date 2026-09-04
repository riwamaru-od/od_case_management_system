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

/**
 * 全案件DB（当期分）から、中止済みを除く全案件を読み出す。
 * 各レポートはこの結果を絞り込んで使う。
 */
function readActiveCasesFromDb_() {
  const sheet = getActiveDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CASE_DATA_START_ROW) return [];
  const values = sheet.getRange(CASE_DATA_START_ROW, 1, lastRow - CASE_DATA_START_ROW + 1, CASE_LAST_COL).getValues();

  const cell = val => (val instanceof Date ? formatDateTime_(val) : (val == null ? '' : String(val)));
  return values
    .filter(row => row[CASE_COLS.CASE_NO - 1] && row[CASE_COLS.STATUS - 1] !== STATUS.CANCELLED)
    .map(row => ({
      caseNo: cell(row[CASE_COLS.CASE_NO - 1]),
      clientName: cell(row[CASE_COLS.CLIENT_NAME - 1]),
      caseName: cell(row[CASE_COLS.CASE_NAME - 1]),
      staffInCharge: cell(row[CASE_COLS.STAFF_IN_CHARGE - 1]),
      billingSchedule: cell(row[CASE_COLS.BILLING_SCHEDULE - 1]),
      status: cell(row[CASE_COLS.STATUS - 1]),
      billingStatus: cell(row[CASE_COLS.BILLING_STATUS - 1]),
      quoteStartedAt: row[CASE_COLS.QUOTE_STARTED_AT - 1],
      invoiceStartedAt: row[CASE_COLS.INVOICE_STARTED_AT - 1],
    }));
}

/** 全案件DB（当期分）から、指定した請求予定ラベルに合致する案件を抽出する */
function getCasesByBillingSchedule_(targetLabels) {
  return readActiveCasesFromDb_().filter(c => targetLabels.indexOf(c.billingSchedule) !== -1);
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

// ------------------------------------------------------------------
// 未承認案件レポート（毎月25日の朝、総務ロール保持者へメール送信）
//   1. 各種書類の未承認一覧（作成中・承認待ちの案件）
//   2. 請求済みだが最終承認が済んでいない案件の一覧
// ------------------------------------------------------------------

/** 未承認レポートの送信日（毎月この日の朝に送信する） */
const UNAPPROVED_REPORT_DAY = 25;

/** 「未承認」とみなすステータス（作成中・承認待ちの両方を対象にする） */
const UNAPPROVED_STATUSES = [
  STATUS.QUOTE_IN_PROGRESS, STATUS.QUOTE_DRAFTED,
  STATUS.INVOICE_IN_PROGRESS, STATUS.INVOICE_DRAFTED,
];

/** 朝の日次処理から呼ばれる。当日が送信日であれば未承認案件レポートを送信する */
function sendUnapprovedSummaryReportIfNeeded_() {
  if (new Date().getDate() !== UNAPPROVED_REPORT_DAY) return;

  const cases = readActiveCasesFromDb_();
  const unapproved = cases.filter(c => UNAPPROVED_STATUSES.indexOf(c.status) !== -1);
  // 請求済み（請求書・納品書のPDF出力まで完了）だが、最終承認がまだの案件
  const billedNotFinalApproved = cases.filter(c => c.billingStatus === BILLING_STATUS.BILLED);

  const msg = buildUnapprovedSummaryMessage_(unapproved, billedNotFinalApproved);
  withLock_('未承認案件レポート送信', () => {
    notifyAdminDept_(msg.subject, msg.body, msg.htmlBody);
    appendOperationLog_('', '未承認案件レポート送信',
      `未承認 ${unapproved.length}件 / 請求済み未承認 ${billedNotFinalApproved.length}件`, false);
  });
}

/** 未承認案件レポートの件名・本文（プレーンテキスト／HTML）を組み立てる */
function buildUnapprovedSummaryMessage_(unapproved, billedNotFinalApproved) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');
  const subject = `【未承認案件レポート】${today}時点`;

  const lines = [
    `${today}時点の未承認案件レポート`,
    '',
    `■ 書類が未承認の案件（${unapproved.length}件）`,
    ...(unapproved.length
      ? unapproved.map(c => `・${c.caseNo}｜${c.clientName}｜${c.caseName}｜担当:${c.staffInCharge}｜${c.status}`)
      : ['該当なし']),
    '',
    `■ 請求済みだが最終承認が未完了の案件（${billedNotFinalApproved.length}件）`,
    ...(billedNotFinalApproved.length
      ? billedNotFinalApproved.map(c => `・${c.caseNo}｜${c.clientName}｜${c.caseName}｜担当:${c.staffInCharge}｜${c.status}`)
      : ['該当なし']),
  ];

  const htmlBody = `
  <div style="font-family:'Hiragino Kaku Gothic ProN','Meiryo',sans-serif;max-width:640px;margin:0 auto;padding:24px;background-color:#f7f7f7;">
    <div style="background-color:#ffffff;border-radius:8px;padding:24px;border:1px solid #e5e5e5;">
      <h2 style="margin:0 0 16px 0;font-size:16px;color:#222222;">${escapeHtmlForMail_(today)}時点の未承認案件レポート</h2>
      ${buildCaseTableHtml_(`書類が未承認の案件（${unapproved.length}件）`, unapproved)}
      ${buildCaseTableHtml_(`請求済みだが最終承認が未完了の案件（${billedNotFinalApproved.length}件）`, billedNotFinalApproved)}
      <p style="margin-top:24px;font-size:11px;color:#999999;">本メールは案件管理・見積書自動作成システムより自動送信されています。</p>
    </div>
  </div>`;

  return { subject, body: lines.join('\n'), htmlBody };
}

/** レポート用の案件一覧テーブル（HTML）。該当が無い場合はその旨を表示する */
function buildCaseTableHtml_(headline, cases) {
  const headerCellStyle = 'padding:6px 10px;text-align:left;font-size:11px;color:#666666;border-bottom:2px solid #dddddd;white-space:nowrap;';
  const cellStyle = 'padding:6px 10px;border-bottom:1px solid #e5e5e5;font-size:12px;color:#222222;';
  const heading = `<h3 style="margin:20px 0 8px 0;font-size:13px;color:#222222;">${escapeHtmlForMail_(headline)}</h3>`;

  if (!cases.length) {
    return `${heading}<p style="margin:0;font-size:12px;color:#666666;">該当なし</p>`;
  }

  const rowsHtml = cases.map(c => `
    <tr>
      <td style="${cellStyle}white-space:nowrap;">${escapeHtmlForMail_(c.caseNo)}</td>
      <td style="${cellStyle}">${escapeHtmlForMail_(c.clientName)}</td>
      <td style="${cellStyle}">${escapeHtmlForMail_(c.caseName)}</td>
      <td style="${cellStyle}white-space:nowrap;">${escapeHtmlForMail_(c.staffInCharge)}</td>
      <td style="${cellStyle}white-space:nowrap;">${escapeHtmlForMail_(c.status)}</td>
    </tr>`).join('');

  return `${heading}
    <table style="width:100%;border-collapse:collapse;">
      <tr>
        <th style="${headerCellStyle}">案件番号</th>
        <th style="${headerCellStyle}">取引先</th>
        <th style="${headerCellStyle}">案件名</th>
        <th style="${headerCellStyle}">担当</th>
        <th style="${headerCellStyle}">ステータス</th>
      </tr>
      ${rowsHtml}
    </table>`;
}

// ------------------------------------------------------------------
// 放置書類アラート（毎朝、作成中のまま一定日数を過ぎた書類を担当者へ通知）
// ------------------------------------------------------------------

/** 「作成中のまま放置されている」とみなす日数 */
const STALLED_DOCUMENT_ALERT_DAYS = 7;

/**
 * 朝の日次処理から呼ばれる。書類を作成（着手）したまま「完成」されずに
 * STALLED_DOCUMENT_ALERT_DAYS 日が過ぎた案件を、案件の担当者へ通知する。
 *
 * 経過日数は着手日時（書類ファイルを作成・再作成した日時）で判定する。
 * 作成日時（QUOTE_CREATED_AT 等）は「完成」ボタン時に記録されるため、
 * 作成中の書類では空欄であり判定に使えない。
 */
function sendStalledDocumentAlertsIfNeeded_() {
  const stalled = findStalledDocuments_();
  if (!stalled.length) return;

  // 担当者ごとにまとめて1通にする
  const byStaffName = {};
  stalled.forEach(item => {
    const key = item.staffInCharge || '';
    if (!byStaffName[key]) byStaffName[key] = [];
    byStaffName[key].push(item);
  });

  withLock_('放置書類アラート送信', () => {
    Object.keys(byStaffName).forEach(staffName => {
      const items = byStaffName[staffName];
      let staff = null;
      try {
        staff = findStaffByName_(staffName);
      } catch (e) {
        console.warn(`放置書類アラートの宛先を解決できませんでした（担当: ${staffName}）: ${e}`);
      }
      if (!staff) return;

      const msg = buildStalledDocumentMessage_(items);
      notifyStaff_(staff.email, msg.subject, msg.body, msg.htmlBody);
    });
    appendOperationLog_('', '放置書類アラート送信', `${stalled.length}件`, false);
  });
}

/** 着手したまま規定日数を過ぎ、まだ「完成」されていない書類を洗い出す */
function findStalledDocuments_() {
  const thresholdMs = STALLED_DOCUMENT_ALERT_DAYS * 24 * 60 * 60 * 1000;
  const now = new Date().getTime();
  const results = [];

  readActiveCasesFromDb_().forEach(c => {
    [
      { label: DOC_TYPES.quote.label, status: STATUS.QUOTE_IN_PROGRESS, startedAt: c.quoteStartedAt },
      { label: DOC_TYPES.invoice.label, status: STATUS.INVOICE_IN_PROGRESS, startedAt: c.invoiceStartedAt },
    ].forEach(doc => {
      if (c.status !== doc.status) return;
      const startedAt = parseStartedAt_(doc.startedAt);
      if (!startedAt) return; // 着手日時が未記録（この機能の導入前に作成された書類）は対象外
      const elapsedDays = Math.floor((now - startedAt.getTime()) / (24 * 60 * 60 * 1000));
      if (now - startedAt.getTime() < thresholdMs) return;
      results.push({
        caseNo: c.caseNo, clientName: c.clientName, caseName: c.caseName,
        staffInCharge: c.staffInCharge, docLabel: doc.label,
        startedAt: formatDateTime_(startedAt), elapsedDays: elapsedDays,
      });
    });
  });

  return results;
}

/** 着手日時セルの値（Date または "yyyy/MM/dd HH:mm" 文字列）を Date に変換する */
function parseStartedAt_(value) {
  if (value instanceof Date) return value;
  const text = String(value == null ? '' : value).trim();
  if (!text) return null;
  const parsed = new Date(text.replace(/-/g, '/'));
  return isNaN(parsed.getTime()) ? null : parsed;
}

/** 放置書類アラートの文面を組み立てる */
function buildStalledDocumentMessage_(items) {
  const subject = `【要対応】作成中のまま${STALLED_DOCUMENT_ALERT_DAYS}日以上経過した書類があります（${items.length}件）`;

  const lines = [
    `作成に着手したまま「完成（承認依頼）」が行われていない書類が${items.length}件あります。`,
    '内容をご確認のうえ、承認依頼または案件中止の操作をお願いします。',
    '',
    ...items.map(i => `・${i.caseNo}｜${i.clientName}｜${i.caseName}｜${i.docLabel}｜着手:${i.startedAt}（${i.elapsedDays}日経過）`),
  ];

  const rows = items.map(i => ({
    label: `${i.caseNo}｜${i.docLabel}`,
    value: `${i.caseName}（${i.clientName}）／着手 ${i.startedAt}・${i.elapsedDays}日経過`,
  }));
  const htmlBody = buildEmailHtml_(
    `作成中のまま${STALLED_DOCUMENT_ALERT_DAYS}日以上経過した書類があります`, rows, null, null);

  return { subject, body: lines.join('\n'), htmlBody };
}
