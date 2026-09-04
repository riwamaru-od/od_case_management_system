/**
 * RoleService.gs
 * 社員DBを参照して、権限判定・通知先（メール／Chatwork Room ID）の取得を行う。
 * ロールはスクリプトプロパティではなく社員DBのロール1〜6列で管理する（確定仕様1章・4章）。
 */

/** 社員DB全行をオブジェクト配列で取得（軽量キャッシュ付き） */
function getAllStaff_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('ALL_STAFF');
  if (cached) return JSON.parse(cached);

  const sheet = getStaffDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, STAFF_COLS.ROLE_6).getValues();
  const staff = values
    .filter(row => row[STAFF_COLS.EMAIL - 1])
    .map(row => ({
      staffNo: row[STAFF_COLS.STAFF_NO - 1],
      name: row[STAFF_COLS.STAFF_NAME - 1],
      email: String(row[STAFF_COLS.EMAIL - 1]).trim().toLowerCase(),
      chatworkRoomId: row[STAFF_COLS.CHATWORK_ROOM_ID - 1],
      department: row[STAFF_COLS.DEPARTMENT - 1],
      roles: STAFF_ROLE_COLS.map(colIndex => row[colIndex - 1]).filter(Boolean),
    }));

  cache.put('ALL_STAFF', JSON.stringify(staff), 300); // 5分キャッシュ
  return staff;
}

/** メールアドレスから社員レコードを取得。見つからなければ null。 */
function findStaffByEmail_(email) {
  const normalized = String(email).trim().toLowerCase();
  return getAllStaff_().find(s => s.email === normalized) || null;
}

/** 指定ロールを1つでも持っているか */
function hasAnyRole_(email, roleList) {
  const staff = findStaffByEmail_(email);
  if (!staff) return false;
  const normalizedTargetRoles = (roleList || []).map(r => String(r).trim().toLowerCase());
  return staff.roles.some(r => normalizedTargetRoles.indexOf(String(r).trim().toLowerCase()) !== -1);
}

/**
 * ロールチェック。権限が無ければ AppError を投げる。
 * サイドバー側では事前にボタンを非活性にするが、サーバー側でも必ず再チェックする。
 */
function assertRole_(email, roleList, actionLabel) {
  if (!hasAnyRole_(email, roleList)) {
    throw AppError_('FORBIDDEN', `この操作（${actionLabel}）を実行する権限がありません。`);
  }
}

function isQuoteApprover_(email) {
  return hasAnyRole_(email, REQUIRED_ROLES.APPROVE_QUOTE);
}
function isInvoiceApprover_(email) {
  return hasAnyRole_(email, REQUIRED_ROLES.APPROVE_INVOICE);
}
function isFinalApprover_(email) {
  return hasAnyRole_(email, REQUIRED_ROLES.FINAL_APPROVE);
}

/** 担当者名（プルダウン選択された社員名）から社員レコードを取得 */
function findStaffByName_(name) {
  const staff = getAllStaff_().find(s => s.name === name);
  if (!staff) {
    throw AppError_('STAFF_NOT_FOUND', `社員DBに「${name}」が見つかりません。`);
  }
  return staff;
}

/** サイドバー用に、社員名の一覧を返す（担当のプルダウン等に利用） */
function listStaffNames_() {
  return getAllStaff_().map(s => s.name);
}

/**
 * 承認依頼先として指定できる社員の一覧を返す（サイドバーのプルダウン用）。
 * 対象は、その書類種別の承認操作が可能なロール（見積書なら「見積書承認」「Admin」「総務」）
 * を1つ以上持ち、メールアドレスが登録されている社員。
 */
function listApproversForDocType_(docTypeKey) {
  const roles = DOC_TYPES[docTypeKey].approverRoles;
  return getAllStaff_()
    .filter(s => s.email && roles.some(role => s.roles.indexOf(role) !== -1))
    .map(s => ({ name: s.name, email: s.email }));
}

/** 総務ロールを持つ全員（承認依頼のメール通知先） */
function getAdminDeptStaff_() {
  return getAllStaff_().filter(s => s.roles.indexOf(ROLES.ADMIN_DEPT) !== -1 || s.roles.indexOf(ROLES.ADMIN) !== -1);
}
