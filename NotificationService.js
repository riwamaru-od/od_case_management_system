/**
 * NotificationService.gs
 * 通知ルール（確定仕様5章）:
 *   - 総務ロールを持つ人 → メール通知（HTML形式。項目を表で整形する）
 *   - それ以外 → Chatwork API 経由でメッセージ通知（Room IDは社員DBから取得。プレーンテキスト）
 */

const CHATWORK_API_BASE = 'https://api.chatwork.com/v2';

/**
 * 指定した社員（メールアドレス）へ、役割に応じた方法で通知する。
 * @param {string} email 通知対象の社員メールアドレス
 * @param {string} subject 件名（メールの場合のみ使用）
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
  if (isAdminDept) {
    sendMailNotification_(staff.email, subject, body, htmlBody);
  } else {
    sendChatworkNotification_(staff.chatworkRoomId, `${subject}\n${body}`);
  }
}

/** 総務ロールを持つ全員に通知（承認依頼など） */
function notifyAdminDept_(subject, body, htmlBody) {
  getAdminDeptStaff_().forEach(staff => sendMailNotification_(staff.email, subject, body, htmlBody));
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

/**
 * 総務向け通知メールの共通HTMLテンプレート。
 * @param {string} headline 見出し
 * @param {Array<{label:string, value:string}>} rows 項目テーブルの行
 * @param {string} [linkUrl] 書類などへのリンクURL（あればボタン風リンクを表示）
 * @param {string} [linkLabel] リンクのラベル文言
 */
function buildEmailHtml_(headline, rows, linkUrl, linkLabel) {
  const rowsHtml = rows.map(r => `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid #e5e5e5;color:#666666;font-size:13px;white-space:nowrap;vertical-align:top;">${escapeHtmlForMail_(r.label)}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #e5e5e5;color:#222222;font-size:13px;">${escapeHtmlForMail_(r.value)}</td>
    </tr>`).join('');

  const buttonHtml = linkUrl ? `
    <div style="margin-top:20px;">
      <a href="${escapeHtmlForMail_(linkUrl)}" style="display:inline-block;padding:10px 20px;background-color:#1a73e8;color:#ffffff;text-decoration:none;border-radius:4px;font-size:14px;">${escapeHtmlForMail_(linkLabel || '書類を開く')}</a>
    </div>` : '';

  return `
  <div style="font-family:'Hiragino Kaku Gothic ProN','Meiryo',sans-serif;max-width:560px;margin:0 auto;padding:24px;background-color:#f7f7f7;">
    <div style="background-color:#ffffff;border-radius:8px;padding:24px;border:1px solid #e5e5e5;">
      <h2 style="margin:0 0 16px 0;font-size:16px;color:#222222;">${escapeHtmlForMail_(headline)}</h2>
      <table style="width:100%;border-collapse:collapse;">${rowsHtml}</table>
      ${buttonHtml}
      <p style="margin-top:24px;font-size:11px;color:#999999;">本メールは案件管理・見積書自動作成システムより自動送信されています。</p>
    </div>
  </div>`;
}

/** 承認依頼通知の文面を組み立てる共通ヘルパー */
function buildApprovalRequestMessage_(docTypeLabel, caseInfo, comment) {
  const subject = `【承認依頼】${caseInfo.caseName}（${docTypeLabel}）`;
  const lines = [
    `案件番号: ${caseInfo.caseNo}`,
    `案件名: ${caseInfo.caseName}`,
    `取引先: ${caseInfo.clientName}`,
    `書類種別: ${docTypeLabel}`,
    `依頼者: ${caseInfo.requesterName}`,
  ];
  const rows = [
    { label: '案件番号', value: caseInfo.caseNo },
    { label: '案件名', value: caseInfo.caseName },
    { label: '取引先', value: caseInfo.clientName },
    { label: '書類種別', value: docTypeLabel },
    { label: '依頼者', value: caseInfo.requesterName },
  ];
  if (comment) {
    lines.push(`コメント: ${comment}`);
    rows.push({ label: 'コメント', value: comment });
  }
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);

  const body = lines.join('\n');
  const htmlBody = buildEmailHtml_(`${docTypeLabel}の承認依頼が届いています`, rows, caseInfo.docUrl, `${docTypeLabel}を確認する`);
  return { subject, body, htmlBody };
}

/** 承認完了通知の文面 */
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
    { label: '結果', value: `${docTypeLabel}が承認されました。` },
  ];
  if (comment) {
    lines.push(`承認コメント: ${comment}`);
    rows.push({ label: '承認コメント', value: comment });
  }
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);

  const body = lines.join('\n');
  const htmlBody = buildEmailHtml_(`${docTypeLabel}が承認されました`, rows, caseInfo.docUrl, `${docTypeLabel}を確認する`);
  return { subject, body, htmlBody };
}
