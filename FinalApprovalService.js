/**
 * FinalApprovalService.gs
 * フェーズ8: 最終承認・クロージング（確定仕様4章・8章に基づく実装）。
 *  - 権限: 総務／Admin／最終承認 のいずれかのロールを持つ人のみ実行可能
 *  - 前提条件: 請求ステータスが「請求済み」（請求書または納品書の印刷／PDF出力が
 *    済んでいる = ApprovalService.markBillingCompletedIfApplicable_ 済み）であること
 *  - 完了後: ステータス／請求ステータスを更新し、全案件DBには残したまま
 *    メイン画面UIから当該行を物理削除する（CaseService.removeCaseFromUiAfterFinalApproval_）
 */

function finalApproveForCase_(caseNo, comment) {
  return withLock_('最終承認', () => {
    const email = getActiveUserEmail_();
    assertRole_(email, REQUIRED_ROLES.FINAL_APPROVE, '最終承認');

    const caseInfo = getCaseInfo_(caseNo);
    if (caseInfo.billingStatus !== BILLING_STATUS.BILLED) {
      throw AppError_(
        'INVALID_STATE',
        '請求書・納品書の印刷またはPDF出力が完了していないため、最終承認できません。'
      );
    }

    const staff = findStaffByEmail_(email);
    const now = new Date();

    setCaseFields_(caseNo, {
      [CASE_COLS.STATUS]: STATUS.FINAL_APPROVED,
      [CASE_COLS.BILLING_STATUS]: BILLING_STATUS.FINAL_APPROVED,
      [CASE_COLS.FINAL_APPROVER]: staff ? staff.name : email,
      [CASE_COLS.FINAL_APPROVED_AT]: formatDateTime_(now),
    });

    // 作成者（見積書作成者を代表として通知）へ完了連絡
    const creatorStaff = getAllStaff_().find(s => s.name === caseInfo.quoteCreator);
    if (creatorStaff) {
      const subject = `【最終承認完了】${caseInfo.caseName}`;
      const lines = [
        `案件番号: ${caseInfo.caseNo}`,
        `案件名: ${caseInfo.caseName}`,
        '最終承認が完了し、案件がクローズされました。',
      ];
      if (comment) lines.push(`コメント: ${comment}`);
      notifyStaff_(creatorStaff.email, subject, lines.join('\n'));
    }

    // 全案件DBへ反映済みの状態で、メイン画面UIから当該行を削除する
    removeCaseFromUiAfterFinalApproval_(caseNo);

    return { status: STATUS.FINAL_APPROVED };
  });
}
