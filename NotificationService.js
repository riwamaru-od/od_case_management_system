/**
 * NotificationService.gs
 * 通知ルール:
 *   社員DBの「メール通知」「Chat通知」チェックボックスに従って配信する。
 *   - メール通知にチェック → メール（HTML形式。項目を表で整形する）
 *   - Chat通知にチェック → Chatwork API 経由のメッセージ（Room IDは社員DBから取得。プレーンテキスト）
 *   - 両方にチェックすれば両方へ、どちらも外せば通知しない
 *   どちらのセルもチェックボックスになっていない（未設定の）社員は、従来のルール
 *   （総務ロール → メール、それ以外 → Chatwork）で送る。
 */

const CHATWORK_API_BASE = 'https://api.chatwork.com/v2';

/**
 * 社員1人へ、社員DBのチェックボックスに従って通知する。
 * @param {Object} staff 社員レコード（getAllStaff_ の要素）
 * @param {boolean} defaultToMail チェックボックスが未設定の場合に、メールで送るか（false なら Chatwork）
 */
function deliverToStaff_(staff, subject, body, htmlBody, defaultToMail) {
  const isUnset = staff.mailNotify === null && staff.chatNotify === null;
  const useMail = isUnset ? defaultToMail : staff.mailNotify === true;
  const useChat = isUnset ? !defaultToMail : staff.chatNotify === true;

  if (useMail) sendMailNotification_(staff.email, subject, body, htmlBody);
  if (useChat) sendChatworkNotification_(staff.chatworkRoomId, `${subject}\n${body}`);
}

/**
 * 指定した社員（メールアドレス）へ通知する。
 * @param {string} email 通知対象の社員メールアドレス
 * @param {string} subject 件名（Chatworkでは本文の1行目になる）
 * @param {string} body 本文（プレーンテキスト。HTML非対応クライアント向けのフォールバックも兼ねる）
 * @param {string} [htmlBody] HTML形式の本文（メール送信時、指定があれば使用）
 */
function notifyStaff_(email, subject, body, htmlBody) {
  const staff = findStaffByEmail_(email);
  if (!staff) {
    console.warn(`notifyStaff_: 社員DBに見つからないため通知をスキップ: ${email}`);
    return;
  }
  const isAdminDept = staff.roles.indexOf(ROLES.ADMIN_DEPT) !== -1;
  deliverToStaff_(staff, subject, body, htmlBody, isAdminDept);
}

/** 総務ロールを持つ全員に通知（定期レポートなど） */
function notifyAdminDept_(subject, body, htmlBody) {
  getAdminDeptStaff_().forEach(staff => deliverToStaff_(staff, subject, body, htmlBody, true));
}

/** 承認依頼の件名に付ける、承認者が指定済みであることを示す前置き */
const APPROVER_DESIGNATED_SUBJECT_PREFIX = '*承認者指定済み* ';

/** 承認依頼の宛先の種類（宛先によって文面を変える） */
const APPROVAL_REQUEST_AUDIENCE = {
  ALL: 'all',               // 承認者の指定なし（総務全員が対応候補）
  DESIGNATED: 'designated', // 指定された承認者本人
  OTHERS: 'others',         // 承認者の指定あり・指定された本人以外の総務（対応不要）
};

/**
 * 承認依頼を通知する。
 * 総務ロール保持者全員へ送るのは従来どおり。加えて承認者が指定されている場合は、
 * 指定された本人へも通知する。総務宛のうち「指定された本人以外」には、件名に
 * 「*承認者指定済み*」を前置きし、本文にも指定された承認者と「対応不要」である旨を記載する。
 * @param {string} docTypeLabel 書類種別の表示名
 * @param {Object} caseInfo 文面に使う案件情報（buildApprovalRequestMessage_ 参照）
 * @param {string} comment 依頼者のコメント
 * @param {string} [designatedEmail] 指定された承認者のメールアドレス（未指定なら従来どおり総務全員のみ）
 */
function notifyApprovalRequest_(docTypeLabel, caseInfo, comment, designatedEmail) {
  const designated = designatedEmail ? findStaffByEmail_(designatedEmail) : null;
  const adminDeptStaff = getAdminDeptStaff_();
  const info = Object.assign({}, caseInfo, { designatedApproverName: designated ? designated.name : '' });

  const build = audience => buildApprovalRequestMessage_(docTypeLabel, info, comment, audience);
  const msgForAll = build(designated ? APPROVAL_REQUEST_AUDIENCE.OTHERS : APPROVAL_REQUEST_AUDIENCE.ALL);
  const msgForDesignated = designated ? build(APPROVAL_REQUEST_AUDIENCE.DESIGNATED) : null;

  adminDeptStaff.forEach(staff => {
    const msg = (designated && staff.email === designated.email) ? msgForDesignated : msgForAll;
    deliverToStaff_(staff, msg.subject, msg.body, msg.htmlBody, true);
  });

  // 指定された承認者が総務以外の場合は、その人の通知手段（Chatwork等）で個別に送る
  if (designated && !adminDeptStaff.some(s => s.email === designated.email)) {
    notifyStaff_(designated.email, msgForDesignated.subject, msgForDesignated.body, msgForDesignated.htmlBody);
  }
}

function sendMailNotification_(email, subject, body, htmlBody) {
  try {
    const options = { to: email, subject: subject, body: body };
    if (htmlBody) options.htmlBody = htmlBody;
    MailApp.sendEmail(options);
  } catch (e) {
    console.error(`メール通知に失敗しました(${email}): ${e}`);
  }
}

function sendChatworkNotification_(roomId, message) {
  if (!roomId) {
    console.warn('sendChatworkNotification_: Chatwork Room ID が未設定のためスキップ');
    return;
  }
  try {
    UrlFetchApp.fetch(`${CHATWORK_API_BASE}/rooms/${roomId}/messages`, {
      method: 'post',
      headers: { 'X-ChatWorkToken': getProp_(PROP_KEYS.CHATWORK_API_TOKEN) },
      payload: { body: message },
      muteHttpExceptions: true,
    });
  } catch (e) {
    console.error(`Chatwork通知に失敗しました(room:${roomId}): ${e}`);
  }
}

/** HTMLメール内で使うためのエスケープ（メールクライアントによる <style> 剥がしを避け、装飾はインラインスタイルで統一する） */
function escapeHtmlForMail_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

/** HTMLメールの配色。通知の種類ごとにアクセント色を変え、一目で内容の性質が分かるようにする */
const EMAIL_TONES = {
  info: { color: '#1a73e8', soft: '#e8f0fe' },    // 依頼・レポート
  success: { color: '#188038', soft: '#e6f4ea' }, // 完了
  danger: { color: '#d93025', soft: '#fce8e6' },  // 差し戻し
  warning: { color: '#e37400', soft: '#fef7e0' }, // 要対応
};
const EMAIL_FONT = "'Hiragino Kaku Gothic ProN','Hiragino Sans','Meiryo','Noto Sans JP',sans-serif";
const EMAIL_BORDER = '#e0e3e7';

/** ステータス文言に応じた色付きバッジ */
function statusBadgeHtml_(text) {
  const t = String(text == null ? '' : text);
  let c = { color: '#5f6368', soft: '#f1f3f4' };
  if (t.indexOf('未請求') !== -1) c = { color: '#c5221f', soft: '#fce8e6' };
  else if (t.indexOf('承認待ち') !== -1) c = { color: '#b06000', soft: '#fef7e0' };
  else if (t.indexOf('作成中') !== -1) c = { color: '#1967d2', soft: '#e8f0fe' };
  else if (t.indexOf('承認済み') !== -1) c = { color: '#137333', soft: '#e6f4ea' };
  else if (t.indexOf('請求済み') !== -1) c = { color: '#007b83', soft: '#e4f7fb' };
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;background-color:${c.soft};color:${c.color};font-size:11px;font-weight:bold;white-space:nowrap;">${escapeHtmlForMail_(t)}</span>`;
}

/** 項目名｜値 の表 */
function buildDetailTableHtml_(rows) {
  if (!rows || !rows.length) return '';
  const last = rows.length - 1;
  const rowsHtml = rows.map((r, i) => {
    const border = i === 0 ? '' : `border-top:1px solid ${EMAIL_BORDER};`;
    const radius = (i === 0 ? 'border-top-left-radius:8px;' : '') + (i === last ? 'border-bottom-left-radius:8px;' : '');
    return `
      <tr>
        <th style="width:112px;padding:11px 14px;background-color:#f8f9fa;text-align:left;font-size:12px;font-weight:normal;color:#5f6368;vertical-align:top;white-space:nowrap;${border}${radius}">${escapeHtmlForMail_(r.label)}</th>
        <td style="padding:11px 14px;font-size:14px;line-height:1.6;color:#202124;${border}">${escapeHtmlForMail_(r.value)}</td>
      </tr>`;
  }).join('');
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;border-collapse:separate;border-spacing:0;border:1px solid ${EMAIL_BORDER};border-radius:8px;">${rowsHtml}</table>`;
}

/** コメントや差し戻し理由などを目立たせる囲み */
function buildNoteHtml_(note, defaultTone) {
  const tone = EMAIL_TONES[note.tone || defaultTone] || EMAIL_TONES.info;
  return `
    <div style="margin-top:16px;padding:12px 16px;background-color:${tone.soft};border-left:4px solid ${tone.color};border-radius:6px;">
      <div style="font-size:12px;font-weight:bold;color:${tone.color};">${escapeHtmlForMail_(note.label)}</div>
      <div style="margin-top:4px;font-size:14px;line-height:1.7;color:#202124;white-space:pre-wrap;">${escapeHtmlForMail_(note.text)}</div>
    </div>`;
}

/** レポート冒頭の件数サマリー */
function buildStatsHtml_(stats, tone) {
  if (!stats || !stats.length) return '';
  const cells = stats.map((s, i) => `
    <td style="padding:0 ${i === stats.length - 1 ? 0 : 8}px 0 0;vertical-align:top;width:${Math.floor(100 / stats.length)}%;">
      <div style="padding:12px 14px;background-color:#f8f9fa;border-radius:8px;">
        <div style="font-size:12px;color:#5f6368;">${escapeHtmlForMail_(s.label)}</div>
        <div style="margin-top:2px;font-size:24px;font-weight:bold;color:${s.color || tone.color};">${escapeHtmlForMail_(s.value)}<span style="margin-left:2px;font-size:12px;font-weight:normal;color:#5f6368;">${escapeHtmlForMail_(s.unit || '件')}</span></div>
      </div>
    </td>`).join('');
  return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;border-collapse:collapse;table-layout:fixed;"><tr>${cells}</tr></table>`;
}

/** 書類などへのリンクボタン（Outlook等でも崩れにくいよう表で組む） */
function buildEmailButtonHtml_(url, label, tone) {
  return `
    <table role="presentation" cellpadding="0" cellspacing="0" style="margin-top:24px;border-collapse:separate;">
      <tr><td style="border-radius:8px;background-color:${tone.color};">
        <a href="${escapeHtmlForMail_(url)}" style="display:inline-block;padding:12px 24px;font-size:14px;font-weight:bold;color:#ffffff;text-decoration:none;border-radius:8px;">${escapeHtmlForMail_(label || '書類を開く')} &rarr;</a>
      </td></tr>
    </table>`;
}

/** 表の上に置く小見出し（件数のバッジ付き） */
function buildEmailSectionTitleHtml_(title, count) {
  const countHtml = count == null ? ''
    : `<span style="display:inline-block;margin-left:6px;padding:1px 8px;border-radius:999px;background-color:#f1f3f4;color:#5f6368;font-size:11px;font-weight:bold;vertical-align:middle;">${escapeHtmlForMail_(count)}件</span>`;
  return `<div style="margin:28px 0 10px 0;font-size:14px;font-weight:bold;color:#202124;">${escapeHtmlForMail_(title)}${countHtml}</div>`;
}

/**
 * 通知メールの共通HTMLテンプレート。
 * @param {Object} opts
 * @param {string} opts.tone 配色（EMAIL_TONES のキー）
 * @param {string} opts.eyebrow 見出しの上に出す通知の種類（例: 承認依頼）
 * @param {string} opts.headline 見出し
 * @param {string} [opts.lead] 見出しの下の説明文
 * @param {{label:string, text:string, tone:string}} [opts.notice] 説明文のすぐ下に出す囲み（対応不要などの重要なお知らせ）
 * @param {Array<{label:string, value:string}>} [opts.stats] 件数サマリー
 * @param {Array<{label:string, value:string}>} [opts.rows] 項目テーブルの行
 * @param {Array<{label:string, text:string, tone:string}>} [opts.notes] コメント等の囲み
 * @param {string} [opts.contentHtml] 上記の後に差し込むHTML（案件一覧の表など）
 * @param {string} [opts.linkUrl] 書類などへのリンクURL（あればボタンを表示）
 * @param {string} [opts.linkLabel] ボタンの文言
 */
function buildEmailHtml_(opts) {
  const tone = EMAIL_TONES[opts.tone] || EMAIL_TONES.info;
  const leadHtml = opts.lead
    ? `<p style="margin:8px 0 0 0;font-size:14px;line-height:1.7;color:#5f6368;">${escapeHtmlForMail_(opts.lead)}</p>` : '';
  const notesHtml = (opts.notes || []).filter(n => n && n.text).map(n => buildNoteHtml_(n, opts.tone)).join('');
  const buttonHtml = opts.linkUrl ? buildEmailButtonHtml_(opts.linkUrl, opts.linkLabel, tone) : '';
  // 案件一覧の表を含むメールは列が多いため、少し広めにする
  const maxWidth = opts.contentHtml ? 820 : 640;

  return `
  <div style="margin:0;padding:24px 12px;background-color:#f1f3f4;font-family:${EMAIL_FONT};">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:${maxWidth}px;margin:0 auto;border-collapse:separate;table-layout:fixed;">
      <tr><td style="padding:28px 28px 28px 28px;background-color:#ffffff;border:1px solid ${EMAIL_BORDER};border-top:4px solid ${tone.color};border-radius:12px;">
        <span style="display:inline-block;padding:3px 10px;border-radius:999px;background-color:${tone.soft};color:${tone.color};font-size:12px;font-weight:bold;">${escapeHtmlForMail_(opts.eyebrow)}</span>
        <h1 style="margin:12px 0 0 0;font-size:20px;line-height:1.4;color:#202124;">${escapeHtmlForMail_(opts.headline)}</h1>
        ${leadHtml}
        ${opts.notice ? buildNoteHtml_(opts.notice, opts.tone) : ''}
        ${buildStatsHtml_(opts.stats, tone)}
        ${buildDetailTableHtml_(opts.rows)}
        ${notesHtml}
        ${opts.contentHtml || ''}
        ${buttonHtml}
      </td></tr>
      <tr><td style="padding:16px 8px 0 8px;text-align:center;font-size:11px;line-height:1.6;color:#80868b;">
        本メールは案件管理・見積書自動作成システムより自動送信されています。
      </td></tr>
    </table>
  </div>`;
}

/**
 * 承認依頼通知の文面を組み立てる共通ヘルパー。
 * @param {string} docTypeLabel 書類種別の表示名
 * @param {Object} caseInfo caseNo, caseName, clientName, requesterName, docUrl, designatedApproverName（指定時のみ）
 * @param {string} comment 依頼者のコメント
 * @param {string} [audience] 宛先の種類（APPROVAL_REQUEST_AUDIENCE。省略時は指定なし）
 */
function buildApprovalRequestMessage_(docTypeLabel, caseInfo, comment, audience) {
  const designatedName = caseInfo.designatedApproverName || '';
  const isOthers = audience === APPROVAL_REQUEST_AUDIENCE.OTHERS && designatedName;
  const isDesignated = audience === APPROVAL_REQUEST_AUDIENCE.DESIGNATED && designatedName;

  const baseSubject = `【承認依頼】${caseInfo.caseName}（${docTypeLabel}）`;
  const subject = isOthers ? APPROVER_DESIGNATED_SUBJECT_PREFIX + baseSubject : baseSubject;
  const othersNotice = `この依頼は承認者に${designatedName}さんが指定されているため、対応は不要です（承認・差し戻しは${designatedName}さんが行います）。`;

  const lines = [];
  if (isOthers) lines.push(`※${othersNotice}`, '');
  if (isDesignated) lines.push('※あなたが承認者に指定されています。', '');
  lines.push(
    `案件番号: ${caseInfo.caseNo}`,
    `案件名: ${caseInfo.caseName}`,
    `取引先: ${caseInfo.clientName}`,
    `書類種別: ${docTypeLabel}`,
    `依頼者: ${caseInfo.requesterName}`,
  );
  if (designatedName) lines.push(`指定された承認者: ${designatedName}`);
  if (comment) lines.push(`コメント: ${comment}`);
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);

  const rows = [
    { label: '案件番号', value: caseInfo.caseNo },
    { label: '案件名', value: caseInfo.caseName },
    { label: '取引先', value: caseInfo.clientName },
    { label: '書類種別', value: docTypeLabel },
    { label: '依頼者', value: caseInfo.requesterName },
  ];
  if (designatedName) rows.push({ label: '指定された承認者', value: designatedName });

  let lead = `${caseInfo.requesterName}さんから承認の依頼が届きました。内容を確認し、承認または差し戻しを行ってください。`;
  if (isDesignated) lead = `${caseInfo.requesterName}さんから、あなたを承認者に指定した承認の依頼が届きました。内容を確認し、承認または差し戻しを行ってください。`;
  if (isOthers) lead = `${caseInfo.requesterName}さんから承認の依頼が届きました。この依頼は${designatedName}さんが対応します。`;

  const body = lines.join('\n');
  const htmlBody = buildEmailHtml_({
    tone: isOthers ? 'success' : 'info',
    eyebrow: isOthers ? '承認依頼（対応不要）' : (isDesignated ? '承認依頼（あなたが承認者）' : '承認依頼'),
    headline: isOthers
      ? `${docTypeLabel}の承認依頼（${designatedName}さんが対応）`
      : `${docTypeLabel}の承認依頼が届いています`,
    lead: lead,
    notice: isOthers ? { label: '総務の皆さまの対応は不要です', text: othersNotice } : null,
    rows: rows,
    notes: [{ label: '依頼者からのコメント', text: comment }],
    linkUrl: caseInfo.docUrl,
    linkLabel: `${docTypeLabel}を確認する`,
  });
  return { subject, body, htmlBody };
}

/** 承認完了通知の文面（依頼者宛。誰が承認したかが分かるよう承認者名を含める） */
function buildApprovedMessage_(docTypeLabel, caseInfo, comment) {
  const subject = `【承認完了】${caseInfo.caseName}（${docTypeLabel}）`;
  const lines = [
    `案件番号: ${caseInfo.caseNo}`,
    `案件名: ${caseInfo.caseName}`,
    `書類種別: ${docTypeLabel}が承認されました。`,
  ];
  const rows = [
    { label: '案件番号', value: caseInfo.caseNo },
    { label: '案件名', value: caseInfo.caseName },
    { label: '書類種別', value: docTypeLabel },
  ];
  if (caseInfo.approverName) {
    lines.push(`承認者: ${caseInfo.approverName}`);
    rows.push({ label: '承認者', value: caseInfo.approverName });
  }
  if (comment) lines.push(`承認コメント: ${comment}`);
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);

  const body = lines.join('\n');
  const htmlBody = buildEmailHtml_({
    tone: 'success',
    eyebrow: '承認完了',
    headline: `${docTypeLabel}が承認されました`,
    lead: caseInfo.approverName
      ? `${caseInfo.approverName}さんが${docTypeLabel}を承認しました。`
      : `${docTypeLabel}が承認されました。`,
    rows: rows,
    notes: [{ label: '承認コメント', text: comment }],
    linkUrl: caseInfo.docUrl,
    linkLabel: `${docTypeLabel}を確認する`,
  });
  return { subject, body, htmlBody };
}

/** 差し戻し通知の文面（作成者宛。修正に取りかかれるよう差し戻し理由と次の操作を含める） */
function buildRejectedMessage_(docTypeLabel, caseInfo, comment) {
  const subject = `【差し戻し】${caseInfo.caseName}（${docTypeLabel}）`;
  const nextAction = 'サイドバーの「再作成」ボタンから修正し、再度「完成」で承認を依頼してください。';
  const lines = [
    `案件番号: ${caseInfo.caseNo}`,
    `案件名: ${caseInfo.caseName}`,
    `書類種別: ${docTypeLabel}が差し戻されました。`,
  ];
  const rows = [
    { label: '案件番号', value: caseInfo.caseNo },
    { label: '案件名', value: caseInfo.caseName },
    { label: '書類種別', value: docTypeLabel },
  ];
  if (caseInfo.rejecterName) {
    lines.push(`差し戻した人: ${caseInfo.rejecterName}`);
    rows.push({ label: '差し戻した人', value: caseInfo.rejecterName });
  }
  if (comment) lines.push(`差し戻し理由: ${comment}`);
  lines.push(`対応: ${nextAction}`);
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);

  const body = lines.join('\n');
  const htmlBody = buildEmailHtml_({
    tone: 'danger',
    eyebrow: '差し戻し',
    headline: `${docTypeLabel}が差し戻されました`,
    lead: caseInfo.rejecterName
      ? `${caseInfo.rejecterName}さんが${docTypeLabel}を差し戻しました。理由を確認し、修正をお願いします。`
      : `${docTypeLabel}が差し戻されました。理由を確認し、修正をお願いします。`,
    rows: rows,
    notes: [
      { label: '差し戻し理由', text: comment },
      { label: '次の対応', text: nextAction, tone: 'info' },
    ],
    linkUrl: caseInfo.docUrl,
    linkLabel: `${docTypeLabel}を確認する`,
  });
  return { subject, body, htmlBody };
}

/** 最終承認完了通知の文面（作成者宛） */
function buildFinalApprovedMessage_(caseInfo, comment) {
  const subject = `【最終承認完了】${caseInfo.caseName}`;
  const lines = [
    `案件番号: ${caseInfo.caseNo}`,
    `案件名: ${caseInfo.caseName}`,
    '最終承認が完了し、案件がクローズされました。',
  ];
  const rows = [
    { label: '案件番号', value: caseInfo.caseNo },
    { label: '案件名', value: caseInfo.caseName },
  ];
  if (caseInfo.approverName) {
    lines.push(`最終承認者: ${caseInfo.approverName}`);
    rows.push({ label: '最終承認者', value: caseInfo.approverName });
  }
  if (comment) lines.push(`コメント: ${comment}`);

  const body = lines.join('\n');
  const htmlBody = buildEmailHtml_({
    tone: 'success',
    eyebrow: '最終承認完了',
    headline: '最終承認が完了しました',
    lead: '最終承認が完了し、この案件はクローズされました。お疲れさまでした。',
    rows: rows,
    notes: [{ label: 'コメント', text: comment }],
  });
  return { subject, body, htmlBody };
}

// ------------------------------------------------------------------
// 【テスト用】通知文面の一括確認
// ------------------------------------------------------------------

/** 通知文面の一括確認で送り先にする社員（社員DBの氏名の一部） */
const NOTIFICATION_SAMPLE_STAFF_NAME = '岩丸';

/**
 * 【テスト用】全パターンの通知文面を、サンプルデータで1人の社員へまとめて送る。
 * 社員DBのチェックボックスに関係なく、メールとChatworkの両方へ送る。
 * 件名の先頭に【テスト】を付け、本物の通知と見分けられるようにする。
 *
 * 末尾を「_」にしていないのは、スクリプトエディタの「実行する関数」プルダウンから
 * 選べるようにするため。実際の案件データは読み書きしない。
 */
function sendAllNotificationSamplesForTest() {
  const staff = getAllStaff_().filter(s => String(s.name).indexOf(NOTIFICATION_SAMPLE_STAFF_NAME) !== -1);
  if (staff.length !== 1) {
    throw new Error(`社員DBで「${NOTIFICATION_SAMPLE_STAFF_NAME}」を含む氏名が1件に絞れません（${staff.length}件）。`);
  }
  const target = staff[0];
  const samples = buildNotificationSamples_();

  samples.forEach(sample => {
    const subject = `【テスト】${sample.msg.subject}`;
    sendMailNotification_(target.email, subject, sample.msg.body, sample.msg.htmlBody);
    sendChatworkNotification_(target.chatworkRoomId, `${subject}\n${sample.msg.body}`);
  });

  const message = [
    `${target.name} さん宛に、${samples.length}パターンの通知を送信しました。`,
    `メール: ${target.email}`,
    `Chatwork: ${target.chatworkRoomId ? `Room ID ${target.chatworkRoomId}` : '社員DBにRoom IDが無いため送信していません'}`,
    '',
    ...samples.map((s, i) => `${i + 1}. ${s.label}`),
  ].join('\n');

  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
  return message;
}

/** 全パターンの通知文面をサンプルデータで組み立てる */
function buildNotificationSamples_() {
  const docUrl = getMainSpreadsheet_().getUrl();
  const caseInfo = {
    caseNo: 'TEST-0001', caseName: 'サンプル案件（通知テスト）', clientName: 'サンプル株式会社',
    requesterName: '作成 太郎', approverName: '承認 花子', rejecterName: '承認 花子', docUrl: docUrl,
  };
  const cases = [
    { caseNo: 'TEST-0001', clientName: 'サンプル株式会社', caseName: 'サンプル案件A', staffInCharge: '担当 一郎',
      billingSchedule: '2026/08 後半', status: STATUS.QUOTE_DRAFTED, billingStatus: BILLING_STATUS.NOT_BILLED, monthsOverdue: 2 },
    { caseNo: 'TEST-0002', clientName: 'テスト商事', caseName: 'サンプル案件B', staffInCharge: '担当 二郎',
      billingSchedule: '2026/09 前半', status: STATUS.INVOICE_IN_PROGRESS, billingStatus: BILLING_STATUS.NOT_BILLED, monthsOverdue: 1 },
  ];
  const billed = [
    { caseNo: 'TEST-0003', clientName: 'サンプル株式会社', caseName: 'サンプル案件C', staffInCharge: '担当 一郎',
      status: STATUS.INVOICE_APPROVED, billingStatus: BILLING_STATUS.BILLED },
  ];
  const stalled = [
    { caseNo: 'TEST-0001', clientName: 'サンプル株式会社', caseName: 'サンプル案件A',
      docLabel: DOC_TYPES.quote.label, startedAt: '2026/09/01 10:00', elapsedDays: 10 },
  ];
  const comment = 'サンプルのコメントです。';
  const nextMonth = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);

  const designatedInfo = Object.assign({}, caseInfo, { designatedApproverName: '承認 花子' });
  return [
    { label: '承認依頼（承認者の指定なし）', msg: buildApprovalRequestMessage_(DOC_TYPES.quote.label, caseInfo, comment) },
    { label: '承認依頼（指定された承認者本人宛）',
      msg: buildApprovalRequestMessage_(DOC_TYPES.quote.label, designatedInfo, comment, APPROVAL_REQUEST_AUDIENCE.DESIGNATED) },
    { label: '承認依頼（承認者指定済み・指定された本人以外の総務宛）',
      msg: buildApprovalRequestMessage_(DOC_TYPES.quote.label, designatedInfo, comment, APPROVAL_REQUEST_AUDIENCE.OTHERS) },
    { label: '承認完了', msg: buildApprovedMessage_(DOC_TYPES.quote.label, caseInfo, comment) },
    { label: '差し戻し', msg: buildRejectedMessage_(DOC_TYPES.quote.label, caseInfo, comment) },
    { label: '最終承認完了', msg: buildFinalApprovedMessage_(caseInfo, comment) },
    { label: '請求予定レポート（対象あり）', msg: buildBillingSummaryMessage_(nextMonth, cases) },
    { label: '請求予定レポート（対象なし）', msg: buildBillingSummaryMessage_(nextMonth, []) },
    { label: '未承認案件レポート', msg: buildUnapprovedSummaryMessage_(cases, billed) },
    { label: '放置書類アラート', msg: buildStalledDocumentMessage_(stalled) },
    { label: '請求漏れリマインド（総務宛）', msg: buildOverdueBillingSummaryMessage_(cases) },
    { label: '請求漏れリマインド（担当者宛）', msg: buildOverdueBillingStaffMessage_(cases) },
  ];
}
