/**
 * DocTypes.gs
 * 見積書・請求書・納品書の「差分」を一箇所にまとめた設定オブジェクト。
 * ApprovalService.gs / DocumentService.gs はこの設定を参照して汎用的に動作する。
 */

const DOC_TYPES = {
  quote: {
    key: 'quote',
    label: '見積書',
    folderKind: 'quote',
    cells: () => QUOTE_TEMPLATE_CELLS,
    getTemplateFileId: () => getTemplateFileId_('quote'),
    col: {
      link: CASE_COLS.QUOTE_LINK,
      outputLink: CASE_COLS.QUOTE_OUTPUT_LINK,
      creator: CASE_COLS.QUOTE_CREATOR,
      createdAt: CASE_COLS.QUOTE_CREATED_AT,
      approver: CASE_COLS.QUOTE_APPROVER,
      approvedAt: CASE_COLS.QUOTE_APPROVED_AT,
      outputBy: CASE_COLS.QUOTE_OUTPUT_BY,
      outputAt: CASE_COLS.QUOTE_OUTPUT_AT,
      rejectedAt: CASE_COLS.QUOTE_REJECTED_AT,
      reapprovalPending: CASE_COLS.QUOTE_REAPPROVAL_PENDING,
      startedAt: CASE_COLS.QUOTE_STARTED_AT,
    },
    status: {
      inProgress: STATUS.QUOTE_IN_PROGRESS,
      drafted: STATUS.QUOTE_DRAFTED,
      approved: STATUS.QUOTE_APPROVED,
    },
    approverRoles: REQUIRED_ROLES.APPROVE_QUOTE,
    hasApprovalStep: true,
    // 承認後のPDF/印刷対象フォルダは常に「未請求案件」
    folderForStage: () => SUBFOLDER.UNBILLED,
  },

  invoice: {
    key: 'invoice',
    label: '請求書',
    folderKind: 'invoice',
    cells: () => INVOICE_TEMPLATE_CELLS,
    getTemplateFileId: () => getTemplateFileId_('invoice'),
    col: {
      link: CASE_COLS.INVOICE_LINK,
      outputLink: CASE_COLS.INVOICE_OUTPUT_LINK,
      creator: CASE_COLS.INVOICE_CREATOR,
      createdAt: CASE_COLS.INVOICE_CREATED_AT,
      approver: CASE_COLS.INVOICE_APPROVER,
      approvedAt: CASE_COLS.INVOICE_APPROVED_AT,
      outputBy: CASE_COLS.INVOICE_OUTPUT_BY,
      outputAt: CASE_COLS.INVOICE_OUTPUT_AT,
      rejectedAt: CASE_COLS.INVOICE_REJECTED_AT,
      reapprovalPending: CASE_COLS.INVOICE_REAPPROVAL_PENDING,
      startedAt: CASE_COLS.INVOICE_STARTED_AT,
    },
    status: {
      inProgress: STATUS.INVOICE_IN_PROGRESS,
      drafted: STATUS.INVOICE_DRAFTED,
      approved: STATUS.INVOICE_APPROVED,
    },
    approverRoles: REQUIRED_ROLES.APPROVE_INVOICE,
    hasApprovalStep: true,
    // 作成時点では「未請求案件」、承認済み以降（PDF出力・印刷）は「請求中案件」に格納する
    folderForStage: (stage) => (stage === 'created' ? SUBFOLDER.UNBILLED : SUBFOLDER.BILLING),
  },

  delivery: {
    key: 'delivery',
    label: '納品書',
    folderKind: 'delivery',
    cells: () => DELIVERY_TEMPLATE_CELLS,
    getTemplateFileId: () => getTemplateFileId_('delivery'),
    col: {
      link: CASE_COLS.DELIVERY_LINK,
      outputLink: CASE_COLS.DELIVERY_OUTPUT_LINK,
      creator: CASE_COLS.DELIVERY_CREATOR,
      createdAt: CASE_COLS.DELIVERY_CREATED_AT,
      // 納品書に承認フローは無い（納品書出力者/日時はテンプレート内セルにのみ記録し、
      // 全案件DBの集計列は確定仕様どおり作成者/作成日時のみ持つ）
    },
    status: {}, // 納品書自体には専用ステータスが無い（案件全体のステータスは変更しない）
    hasApprovalStep: false,
    folderForStage: () => null, // xx期フォルダ直下（サブフォルダ無し）
  },
};
