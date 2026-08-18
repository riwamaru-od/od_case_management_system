/**
 * NotificationService.gs
 * 通知ルール（確定仕様5章）:
 *   - 総務ロールを持つ人 → メール通知
 *   - それ以外 → Chatwork API 経由でメッセージ通知（Room IDは社員DBから取得）
 */

const CHATWORK_API_BASE = 'https://api.chatwork.com/v2';

/**
 * 指定した社員（メールアドレス）へ、役割に応じた方法で通知する。
 * @param {string} email 通知対象の社員メールアドレス
 * @param {string} subject 件名（メールの場合のみ使用）
 * @param {string} body 本文
 */
function notifyStaff_(email, subject, body) {
  const staff = findStaffByEmail_(email);
  if (!staff) {
    console.warn(`notifyStaff_: 社員DBに見つからないため通知をスキップ: ${email}`);
    return;
  }
  const isAdminDept = staff.roles.indexOf(ROLES.ADMIN_DEPT) !== -1;
  if (isAdminDept) {
    sendMailNotification_(staff.email, subject, body);
  } else {
    sendChatworkNotification_(staff.chatworkRoomId, `${subject}\n${body}`);
  }
}

/** 総務ロールを持つ全員に通知（承認依頼など） */
function notifyAdminDept_(subject, body) {
  getAdminDeptStaff_().forEach(staff => sendMailNotification_(staff.email, subject, body));
}

function sendMailNotification_(email, subject, body) {
  try {
    MailApp.sendEmail({ to: email, subject: subject, body: body });
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
  if (comment) lines.push(`コメント: ${comment}`);
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);
  return { subject, body: lines.join('\n') };
}

/** 承認完了通知の文面 */
function buildApprovedMessage_(docTypeLabel, caseInfo, comment) {
  const subject = `【承認完了】${caseInfo.caseName}（${docTypeLabel}）`;
  const lines = [
    `案件番号: ${caseInfo.caseNo}`,
    `案件名: ${caseInfo.caseName}`,
    `書類種別: ${docTypeLabel}が承認されました。`,
  ];
  if (comment) lines.push(`承認コメント: ${comment}`);
  if (caseInfo.docUrl) lines.push(`リンク: ${caseInfo.docUrl}`);
  return { subject, body: lines.join('\n') };
}
