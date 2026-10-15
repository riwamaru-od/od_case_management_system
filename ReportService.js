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

  const lines = cases.length
    ? [
      `${monthLabel}分の請求予定案件一覧（${cases.length}件）`, '',
      ...cases.map(c => `・${c.caseNo}｜${c.clientName}｜${c.caseName}｜請求予定:${c.billingSchedule}｜ステータス:${c.status}｜請求ステータス:${c.billingStatus}`),
    ]
    : [`${monthLabel}分の請求予定案件はありません。`];

  const notBilled = cases.filter(c => c.billingStatus === BILLING_STATUS.NOT_BILLED).length;
  const htmlBody = buildEmailHtml_({
    tone: 'info',
    eyebrow: '請求予定レポート',
    headline: `${monthLabel}分の請求予定`,
    lead: `請求予定が${monthLabel}の前半・後半になっている案件の一覧です。`,
    stats: [
      { label: '対象の案件', value: cases.length },
      { label: 'うち未請求', value: notBilled, color: notBilled ? EMAIL_TONES.danger.color : null },
    ],
    contentHtml: buildCaseTableHtml_('対象の案件', cases, [
      { label: '案件番号', value: c => c.caseNo, nowrap: true },
      { label: '取引先', value: c => c.clientName },
      { label: '案件名', value: c => c.caseName },
      { label: '請求予定', value: c => c.billingSchedule, nowrap: true },
      { label: 'ステータス', value: c => c.status, badge: true },
      { label: '請求', value: c => c.billingStatus, badge: true },
    ]),
  });
  return { subject, body: lines.join('\n'), htmlBody };
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
  sendUnapprovedSummaryReport_();
}

/** 未承認案件レポートを送信する（送信日かどうかは判定しない） */
function sendUnapprovedSummaryReport_() {
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

  const htmlBody = buildEmailHtml_({
    tone: 'info',
    eyebrow: '未承認案件レポート',
    headline: `${today}時点の未承認案件`,
    lead: '承認が済んでいない書類と、請求済みで最終承認を待っている案件の一覧です。',
    stats: [
      { label: '書類が未承認', value: unapproved.length },
      { label: '最終承認待ち', value: billedNotFinalApproved.length },
    ],
    contentHtml:
      buildCaseTableHtml_('書類が未承認の案件', unapproved)
      + buildCaseTableHtml_('請求済みだが最終承認が未完了の案件', billedNotFinalApproved),
  });

  return { subject, body: lines.join('\n'), htmlBody };
}

/** レポート用テーブルの既定の列構成 */
function defaultCaseTableColumns_() {
  return [
    { label: '案件番号', value: c => c.caseNo, nowrap: true },
    { label: '取引先', value: c => c.clientName },
    { label: '案件名', value: c => c.caseName },
    { label: '担当', value: c => c.staffInCharge, nowrap: true },
    { label: 'ステータス', value: c => c.status, badge: true },
  ];
}

/**
 * レポート用の案件一覧テーブル（HTML）。見出し行と縞模様つきで、該当が無い場合はその旨を表示する。
 * @param {string} title 表の見出し（件数は自動で添える）
 * @param {Array} cases 案件の配列
 * @param {Array<{label:string, value:function, nowrap:boolean, badge:boolean, align:string}>} [columns]
 *   列構成（省略時は既定）。badge は値をステータスのバッジで、align は 'right' で右寄せにする
 */
function buildCaseTableHtml_(title, cases, columns) {
  const cols = columns || defaultCaseTableColumns_();
  const heading = buildEmailSectionTitleHtml_(title, cases.length);

  if (!cases.length) {
    return `${heading}<div style="padding:16px;border:1px dashed ${EMAIL_BORDER};border-radius:8px;text-align:center;font-size:13px;color:#80868b;">該当する案件はありません</div>`;
  }

  const headerHtml = cols.map(col =>
    `<th style="padding:9px 12px;background-color:#f8f9fa;border-bottom:1px solid ${EMAIL_BORDER};text-align:${col.align || 'left'};font-size:11px;font-weight:bold;color:#5f6368;white-space:nowrap;">${escapeHtmlForMail_(col.label)}</th>`
  ).join('');
  const rowsHtml = cases.map((c, i) => {
    const bg = i % 2 === 1 ? 'background-color:#fafbfc;' : '';
    const border = i === cases.length - 1 ? '' : `border-bottom:1px solid #f1f3f4;`;
    return `<tr>${cols.map(col => {
      const value = col.value(c);
      const inner = col.badge ? statusBadgeHtml_(value) : escapeHtmlForMail_(value);
      // 折り返す列は、1文字ずつ縦に崩れないよう最小幅を持たせる
      const widthStyle = (col.nowrap || col.badge) ? 'white-space:nowrap;' : 'min-width:96px;';
      return `<td style="padding:10px 12px;${bg}${border}text-align:${col.align || 'left'};font-size:13px;line-height:1.5;color:#202124;vertical-align:top;${widthStyle}">${inner}</td>`;
    }).join('')}</tr>`;
  }).join('');

  return `${heading}
    <div style="overflow-x:auto;border:1px solid ${EMAIL_BORDER};border-radius:8px;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
        <tr>${headerHtml}</tr>
        ${rowsHtml}
      </table>
    </div>`;
}

/**
 * 案件を担当者ごとにまとめ、担当者本人へ通知する。
 * 通知手段は notifyStaff_ に従う（総務ロールはメール、それ以外はChatwork）。
 * @param {Array} items staffInCharge を持つ要素の配列
 * @param {function(Array): {subject:string, body:string, htmlBody:string}} buildMessage
 *   その担当者ぶんの要素を受け取り、通知文面を組み立てる関数
 * @return {number} 実際に通知した担当者の人数
 */
function notifyStaffInChargeGrouped_(items, buildMessage) {
  const byStaffName = {};
  items.forEach(item => {
    const key = item.staffInCharge || '';
    if (!byStaffName[key]) byStaffName[key] = [];
    byStaffName[key].push(item);
  });

  let notified = 0;
  Object.keys(byStaffName).forEach(staffName => {
    let staff = null;
    try {
      staff = findStaffByName_(staffName);
    } catch (e) {
      // 担当が未入力、または社員DBに無い名前の場合。全体の送信は止めない
      console.warn(`通知先の担当者を解決できませんでした（担当: ${staffName}）: ${e}`);
    }
    if (!staff) return;
    const msg = buildMessage(byStaffName[staffName]);
    notifyStaff_(staff.email, msg.subject, msg.body, msg.htmlBody);
    notified++;
  });
  return notified;
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

  withLock_('放置書類アラート送信', () => {
    // 担当者ごとにまとめて1通にする
    const notified = notifyStaffInChargeGrouped_(stalled, buildStalledDocumentMessage_);
    appendOperationLog_('', '放置書類アラート送信', `${stalled.length}件 / 担当者${notified}名へ送信`, false);
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
  const request = '内容をご確認のうえ、承認依頼または案件中止の操作をお願いします。';

  const lines = [
    `作成に着手したまま「完成（承認依頼）」が行われていない書類が${items.length}件あります。`,
    request,
    '',
    ...items.map(i => `・${i.caseNo}｜${i.clientName}｜${i.caseName}｜${i.docLabel}｜着手:${i.startedAt}（${i.elapsedDays}日経過）`),
  ];

  const htmlBody = buildEmailHtml_({
    tone: 'warning',
    eyebrow: '要対応',
    headline: `作成中のまま${STALLED_DOCUMENT_ALERT_DAYS}日以上経過した書類があります`,
    lead: `作成に着手したまま「完成（承認依頼）」が行われていない書類です。${request}`,
    contentHtml: buildCaseTableHtml_('対象の書類', items, [
      { label: '案件番号', value: i => i.caseNo, nowrap: true },
      { label: '取引先', value: i => i.clientName },
      { label: '案件名', value: i => i.caseName },
      { label: '書類', value: i => i.docLabel, nowrap: true },
      { label: '着手日時', value: i => i.startedAt, nowrap: true },
      { label: '経過', value: i => `${i.elapsedDays}日`, nowrap: true, align: 'right' },
    ]),
  });

  return { subject, body: lines.join('\n'), htmlBody };
}

// ------------------------------------------------------------------
// 請求漏れリマインド（毎月5日の朝、総務と案件の担当者へ）
//   請求予定が前月以前なのに、請求ステータスが「未請求」のままの案件を通知する。
// ------------------------------------------------------------------

/** 請求漏れリマインドの送信日（毎月この日の朝に送信する） */
const OVERDUE_BILLING_REMINDER_DAY = 5;

/**
 * 「YYYY/MM 前半」「YYYY/MM 後半」形式の請求予定から年月を取り出す。
 * 「未定」「営業案件」など月に紐づかない選択肢や未入力の場合は null を返す
 * （請求予定日が決まっていないため、期限超過の判定ができない）。
 */
function parseBillingScheduleMonth_(label) {
  const matched = String(label == null ? '' : label).trim().match(/^(\d{4})\/(\d{1,2})/);
  if (!matched) return null;
  return { year: Number(matched[1]), month: Number(matched[2]) };
}

/**
 * 朝の日次処理から呼ばれる。当日が送信日であれば、請求予定を過ぎたまま
 * 未請求になっている案件を、総務と案件の担当者の双方へ通知する。
 * 該当が無い月は何も送らない。
 */
function sendOverdueBillingRemindersIfNeeded_() {
  if (new Date().getDate() !== OVERDUE_BILLING_REMINDER_DAY) return;
  sendOverdueBillingReminders_();
}

/** 請求漏れリマインドを送信する（送信日かどうかは判定しない） */
function sendOverdueBillingReminders_() {
  const overdue = findOverdueBillingCases_(new Date());
  if (!overdue.length) return;

  withLock_('請求漏れリマインド送信', () => {
    // 1. 総務ロール保持者全員へ、全件の一覧をメールで送る
    const adminMsg = buildOverdueBillingSummaryMessage_(overdue);
    notifyAdminDept_(adminMsg.subject, adminMsg.body, adminMsg.htmlBody);

    // 2. 案件の担当者本人へ、自分の担当ぶんだけを送る
    const notified = notifyStaffInChargeGrouped_(overdue, buildOverdueBillingStaffMessage_);

    appendOperationLog_('', '請求漏れリマインド送信',
      `${overdue.length}件 / 総務・担当者${notified}名へ送信`, false);
  });
}

/**
 * 請求予定が前月以前で、請求ステータスが「未請求」のままの案件を洗い出す。
 * 当月ぶんはまだ請求期限内のため対象に含めない。
 */
function findOverdueBillingCases_(today) {
  const currentMonthIndex = today.getFullYear() * 12 + today.getMonth(); // getMonth() は0始まり

  return readActiveCasesFromDb_()
    .filter(c => {
      if (c.billingStatus !== BILLING_STATUS.NOT_BILLED) return false;
      const scheduled = parseBillingScheduleMonth_(c.billingSchedule);
      if (!scheduled) return false;
      return (scheduled.year * 12 + (scheduled.month - 1)) < currentMonthIndex;
    })
    .map(c => {
      const scheduled = parseBillingScheduleMonth_(c.billingSchedule);
      const monthsOverdue = currentMonthIndex - (scheduled.year * 12 + (scheduled.month - 1));
      return Object.assign({}, c, { monthsOverdue: monthsOverdue });
    })
    .sort((a, b) => b.monthsOverdue - a.monthsOverdue); // 遅れが大きいものから並べる
}

/** 請求漏れリマインドの列構成（請求予定と遅れ月数を見せる） */
function overdueBillingTableColumns_() {
  return [
    { label: '案件番号', value: c => c.caseNo, nowrap: true },
    { label: '取引先', value: c => c.clientName },
    { label: '案件名', value: c => c.caseName },
    { label: '担当', value: c => c.staffInCharge, nowrap: true },
    { label: '請求予定', value: c => c.billingSchedule, nowrap: true },
    { label: '遅れ', value: c => `${c.monthsOverdue}ヶ月`, nowrap: true, align: 'right' },
    { label: 'ステータス', value: c => c.status, badge: true },
  ];
}

/** 請求漏れリマインドの1件ぶんの行（プレーンテキスト用） */
function overdueBillingTextLine_(c) {
  return `・${c.caseNo}｜${c.clientName}｜${c.caseName}｜担当:${c.staffInCharge}`
    + `｜請求予定:${c.billingSchedule}（${c.monthsOverdue}ヶ月遅れ）｜ステータス:${c.status}`;
}

/** 総務宛（全件一覧）の文面 */
function buildOverdueBillingSummaryMessage_(cases) {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');
  const subject = `【請求漏れリマインド】請求予定を過ぎた未請求の案件が${cases.length}件あります`;
  const request = '請求手続きまたは請求予定の見直しをお願いします。';

  const body = [
    `${today}時点で、請求予定の月を過ぎたまま請求ステータスが「未請求」の案件が${cases.length}件あります。`,
    `内容をご確認のうえ、${request}`,
    '',
    ...cases.map(overdueBillingTextLine_),
  ].join('\n');

  const maxOverdue = cases.reduce((max, c) => Math.max(max, c.monthsOverdue || 0), 0);
  const htmlBody = buildEmailHtml_({
    tone: 'warning',
    eyebrow: '請求漏れリマインド',
    headline: '請求予定を過ぎた未請求の案件があります',
    lead: `${today}時点で、請求予定の月を過ぎたまま「未請求」の案件です。${request}`,
    stats: [
      { label: '対象の案件', value: cases.length },
      { label: '最大の遅れ', value: maxOverdue, unit: 'ヶ月', color: EMAIL_TONES.danger.color },
    ],
    contentHtml: buildCaseTableHtml_('対象の案件', cases, overdueBillingTableColumns_()),
  });

  return { subject, body, htmlBody };
}

/** 担当者宛（自分の担当ぶんのみ）の文面 */
function buildOverdueBillingStaffMessage_(cases) {
  const subject = `【要対応】請求予定を過ぎた未請求の案件があります（${cases.length}件）`;
  const request = '請求手続きを進めるか、請求予定の見直しをお願いします。';

  const body = [
    `ご担当の案件のうち、請求予定の月を過ぎたまま請求ステータスが「未請求」のものが${cases.length}件あります。`,
    request,
    '',
    ...cases.map(overdueBillingTextLine_),
  ].join('\n');

  // 本人宛のため「担当」列は省く
  const columns = overdueBillingTableColumns_().filter(col => col.label !== '担当');
  const htmlBody = buildEmailHtml_({
    tone: 'warning',
    eyebrow: '要対応',
    headline: '請求予定を過ぎた未請求の案件があります',
    lead: `ご担当の案件のうち、請求予定の月を過ぎたまま「未請求」のものです。${request}`,
    contentHtml: buildCaseTableHtml_('対象の案件', cases, columns),
  });

  return { subject, body, htmlBody };
}

/**
 * 【テスト用】定期通知3種を、送信日の判定を無視してその場で送信する。
 * 未承認案件レポート（本来は毎月25日）・請求漏れリマインド（本来は毎月5日）は
 * 送信日でないと動作を確認できないため、動作確認用にこの入口を用意している。
 *
 * 末尾を「_」にしていないのは、スクリプトエディタの「実行する関数」プルダウンから
 * 選べるようにするため（末尾が「_」の関数は表示されない）。
 *
 * 注意: 実際に総務・担当者へメールやChatworkが送信される。
 *      動作確認以外の目的では実行しないこと。
 */
function runScheduledNotificationsForTest() {
  const results = [];
  [
    ['未承認案件レポート', sendUnapprovedSummaryReport_],
    ['請求漏れリマインド', sendOverdueBillingReminders_],
    ['放置書類アラート', sendStalledDocumentAlertsIfNeeded_],
  ].forEach(([label, fn]) => {
    try {
      fn();
      results.push(`○ ${label}: 実行しました`);
    } catch (e) {
      results.push(`× ${label}: ${e && e.message ? e.message : e}`);
    }
  });

  const message = [
    '定期通知のテスト送信を実行しました。',
    '（該当案件が無い通知は、何も送信されずに終了します）',
    '',
    ...results,
  ].join('\n');

  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
  return message;
}
