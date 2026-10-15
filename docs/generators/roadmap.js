/**
 * roadmap.js
 * 導入ロードマップ・タスク一覧。
 * 他のマニュアル（common.js）とは別に、社内標準デザインルール
 * （濃紺 #172145 ＋ ブランドレッド #D43C2B、BIZ UDPゴシック、A4 余白 上30・下25・左右25mm）で組む。
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, Header, Footer, ImageRun, PageNumber, TabStopType,
  LevelFormat,
} = require('docx');

const FONT = 'BIZ UDPゴシック';
const CODE_FONT = 'Consolas';
const mm = v => Math.round(v * 56.6929); // mm → DXA
const W = mm(160); // 本文幅

const COLOR = {
  main: '172145',
  accent: 'D43C2B',
  text: '222222',
  note: '757575',
  rule: 'D5D5D5',
  surface: 'F0F0F0',
  white: 'FFFFFF',
};

const DOC_TITLE = '導入ロードマップ・タスク一覧';

// ---------- 部品 ----------

/** `code`・**太字**・【要確認…】を TextRun に分解する */
function runs(text, base = {}) {
  return String(text).split(/(\*\*[^*]+\*\*|`[^`]+`|【要確認[^】]*】)/).filter(Boolean).map(part => {
    if (part.startsWith('**')) return new TextRun({ ...base, text: part.slice(2, -2), bold: true });
    // 和文はゴシックのまま、英数字だけ等幅にする
    if (part.startsWith('`')) return new TextRun({ ...base, text: part.slice(1, -1), font: { ascii: CODE_FONT, hAnsi: CODE_FONT, eastAsia: FONT } });
    if (part.startsWith('【要確認')) return new TextRun({ ...base, text: part, bold: true, color: COLOR.note });
    return new TextRun({ ...base, text: part });
  });
}

const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const p = (t, opts = {}) => new Paragraph({ children: runs(t, opts.run), spacing: { after: 120 }, ...opts.para });
const note = t => p(t, { run: { size: 18, color: COLOR.note } });
const bullet = t => new Paragraph({ children: runs(t), numbering: { reference: 'bullets', level: 0 }, spacing: { after: 60 } });
const period = t => p(`期間（目安）：${t}`, { run: { color: COLOR.note }, para: { keepNext: true } });

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const RULE = { style: BorderStyle.SINGLE, size: 4, color: COLOR.rule };

/** セル内容: 文字列、または [本文, 補足1, 補足2…]（補足は字下げした「・」行） */
function cellParas(content, opts = {}) {
  const lines = Array.isArray(content) ? content : [content];
  return lines.map((line, i) => new Paragraph({
    children: runs(i === 0 ? line : `・${line}`, { size: 19, color: opts.color, bold: opts.bold }),
    alignment: opts.alignment,
    indent: i === 0 ? undefined : { left: 200 },
    spacing: { after: 0, line: 300 },
  }));
}

function cell(content, width, opts = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 110, right: 110 },
    children: cellParas(content, opts),
  });
}

/** 表: ヘッダーは濃紺＋白文字、罫線は横線のみ */
function table(headers, rows, widths, opts = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const center = opts.center || [];
  return new Table({
    columnWidths: widths,
    width: { size: total, type: WidthType.DXA },
    borders: { top: RULE, bottom: RULE, insideHorizontal: RULE, left: NO_BORDER, right: NO_BORDER, insideVertical: NO_BORDER },
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => cell(h, widths[i], {
          fill: COLOR.main, color: COLOR.white, bold: true,
          alignment: center.includes(i) ? AlignmentType.CENTER : undefined,
        })),
      }),
      ...rows.map(r => new TableRow({
        cantSplit: true,
        children: r.map((c, i) => cell(c, widths[i], {
          alignment: center.includes(i) ? AlignmentType.CENTER : undefined,
        })),
      })),
    ],
  });
}

/** タスク表: [担当, 内容, 済み?] の配列。先頭にチェック欄を付ける（済みなら ☑） */
const TASK_W = [mm(10), mm(24), W - mm(10) - mm(24)];
const tasks = rows => table(['済', '担当', 'タスク'], rows.map(([who, what, isDone]) => [isDone ? '☑' : '☐', who, what]), TASK_W, { center: [0] });

/** ペールグレーの囲み */
function callout(title, lines) {
  return new Table({
    columnWidths: [W],
    width: { size: W, type: WidthType.DXA },
    borders: { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER, insideHorizontal: NO_BORDER, insideVertical: NO_BORDER },
    rows: [new TableRow({
      cantSplit: true,
      children: [new TableCell({
        width: { size: W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: COLOR.surface, color: 'auto' },
        margins: { top: 140, bottom: 140, left: 200, right: 200 },
        children: [
          ...(title ? [new Paragraph({ children: runs(title, { bold: true, color: COLOR.main }), spacing: { after: 60 } })] : []),
          ...lines.map(l => new Paragraph({ children: runs(l), spacing: { after: 40, line: 320 } })),
        ],
      })],
    })],
  });
}

const done = t => callout('完了条件', [t]);
const gap = (after = 120) => new Paragraph({ children: [], spacing: { after } });

// ---------- 本文 ----------

const c = [];

c.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun(DOC_TITLE)] }));
c.push(p('案件管理・見積書自動作成システム', { run: { size: 24, color: COLOR.main, bold: true }, para: { spacing: { after: 80 } } }));
c.push(new Paragraph({
  children: [
    ...runs('作成日：2026/10/13 ｜ 最終更新：2026/10/15 ｜ 作成者：【要確認：部署・氏名】', { color: COLOR.note, size: 19 }),
  ],
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR.main, space: 6 } },
  spacing: { after: 240 },
}));

c.push(h1('概要'));
c.push(table(['項目', '内容'], [
  ['目的', 'テストから本稼働までの作業と順番、各フェーズの完了条件を関係者でそろえる'],
  ['結論', '2026/12/1 の本稼働を目標に、8つのフェーズで進める。各フェーズの完了条件を満たしてから次へ進む'],
  ['お願い', '決裁者は 2026/10/16 までに、本稼働日・進行中案件の扱い・スクリプトの編集権限の扱いの3点を決める'],
], [mm(24), W - mm(24)]));
c.push(gap(80));
c.push(new Paragraph({
  children: [
    new TextRun({ text: '要判断：', bold: true, color: COLOR.accent }),
    ...runs('スプレッドシートの編集者は、管理用アカウントの権限で動く処理を書き換えられます（「リスク・懸念事項」の2）。本稼働前に扱いを決めてください。Webアプリ入口のなりすましは 2026/10/15 に対策済みです（同1）。'),
  ],
  spacing: { after: 120 },
}));
c.push(note('日付はすべて目安です。担当の略記：開発＝開発担当／管理＝管理用アカウントの運用者／利用＝利用者（社員）／決裁＝導入の責任者'));

c.push(h1('目次'));
[
  '全体のスケジュール',
  'フェーズ0　現状整理・決めごと',
  'フェーズ1　開発環境での最終確認',
  'フェーズ2　本番環境の構築',
  'フェーズ3　ユーザーテスト',
  'フェーズ4　不具合修正・再テスト',
  'フェーズ5　本番切り替えの準備',
  'フェーズ6　本稼働・初期サポート',
  'フェーズ7　年間の節目の確認',
  'リスク・懸念事項',
  '関連ドキュメント',
  '改訂履歴',
].forEach(t => c.push(bullet(t)));

// 全体
c.push(h1('全体のスケジュール'));
c.push(table(['#', 'フェーズ', '期間（目安）', 'ゴール'], [
  ['0', '現状整理・決めごと', '〜10/16', '未決事項がなくなり、日程が確定している'],
  ['1', '開発環境での最終確認', '10/19〜10/23', '9月17日修正を含め、開発環境で全機能が動く'],
  ['2', '本番環境の構築', '10/26〜10/30', '本番環境が初期設定マニュアルどおりに構築されている'],
  ['3', 'ユーザーテスト', '11/2〜11/13', 'ユーザーテスト・チェックリスト（L区分を除く188項目）を実施済み'],
  ['4', '不具合修正・再テスト', '11/16〜11/20', 'NG項目がすべて解消し、総合判定「可」が出ている'],
  ['5', '本番切り替えの準備', '11/24〜11/27', 'テストデータが消え、利用者が使い始められる'],
  ['6', '本稼働・初期サポート', '12/1〜12月末', '1か月分の書類作成・請求サイクルが問題なく回る'],
  ['7', '年間の節目の確認', '2027/4/1・5/1・アーカイブ時', '期の切り替えとアーカイブが自動で回る'],
], [mm(8), mm(38), mm(44), W - mm(90)], { center: [0] }));
c.push(gap(80));
c.push(p('本稼働を月初（12/1）にしているのは、請求予定レポート・請求漏れリマインドなど月単位の処理を、1か月まるごと確認できるためです。'));

// フェーズ0
c.push(h1('フェーズ0　現状整理・決めごと'));
c.push(period('〜2026/10/16'));
c.push(h2('決めること'));
c.push(tasks([
  ['決裁', '本稼働日を決める（案：2026/12/1）'],
  ['決裁', ['稼働開始時点で進行中の案件をどう扱うか決める',
    '案A：新規案件からのみシステムで管理する（既存案件は従来の方法で完了させる）',
    '案B：進行中の案件もシステムへ移す。移行機能で旧一覧の案件番号をそのまま引き継げる（完了・中止の案件も履歴として全案件DBに残せる）。書類はシステム上で作り直す']],
  ['決裁', ['ユーザーテストの参加者を決める（最低限、下記のロールを1人ずつ）',
    '管理用アカウント／見積書承認者／請求書承認者／最終承認者／総務（メール通知の受け手）／一般社員']],
  ['決裁・開発', 'スクリプトの編集権限の扱いを決める（「リスク・懸念事項」の2）'],
]));
c.push(h2('確認・準備すること'));
c.push(tasks([
  ['開発', '9月17日修正の動作確認（`9月17日_修正点と動作確認チェックリスト.txt`）がどこまで済んでいるか確認する'],
  ['管理', '本番用の管理用アカウント（専用Googleアカウント）を用意する'],
  ['管理', 'Chatwork APIトークンを取得する（通知を送る側のアカウント）'],
  ['管理', '社員DBに載せる全社員の「氏名・メール・ロール・Chatwork ID・通知方法（メール／Chatwork）」を集める'],
  ['管理', '案Bの場合、旧一覧から移す案件を「案件番号・取引先名・案件名・終了予定・請求予定・担当・移行区分（進行中／完了／中止）」の表にまとめる'],
  ['管理', '取引先DBに載せる取引先（社名・担当者・留意点など）を集める。同じ社名で担当者が異なる取引先を洗い出す'],
  ['管理', '社印画像（背景透過のPNG）を用意する'],
  ['管理', '見積書・請求書・納品書の本番用テンプレートの最終デザインを確定する'],
]));
c.push(gap());
c.push(done('上の決めごとがすべて決まり、本稼働日と各フェーズの日程が関係者に共有されている。'));

// フェーズ1
c.push(h1('フェーズ1　開発環境での最終確認'));
c.push(period('2026/10/19〜10/23'));
c.push(tasks([
  ['開発', '最新コードを開発環境へ反映し、Webアプリを新バージョンで再デプロイする（2026/10/15 v23）', true],
  ['開発', 'リポジトリの `.clasp.json` を、実際に使っている開発環境（`1cHuC…`）に向け直す（2026/10/15）', true],
  ['開発', ['`9月17日_修正点と動作確認チェックリスト.txt` を最後まで実施する',
    '0. 事前作業（列の拡張 36→37、保護範囲をAK列まで、テンプレートの転記先セル変更、取引先プルダウン生成）',
    '1. 請求後の再見積もり（1-1〜1-15）',
    '2. テンプレートの転記位置（2-1〜2-10）',
    '3. 取引先名のプルダウン（3-1〜3-7）',
    '4. 既存機能のデグレード確認（4-1〜4-5）']],
  ['開発', '`9月4日_修正点と動作確認チェックリスト.txt` で未確認の項目があれば実施する'],
  ['開発', 'NGが出た項目を修正し、該当項目を再確認する'],
  ['開発', '仕様書・各マニュアル・ユーザーテスト・チェックリストが最新の実装と一致しているか確認する（直す場合は `docs/generators/` を修正して `npm run build`）'],
  ['開発', ['Webアプリ入口の本人確認とサイドバーの修正（2026/10/15 実装）を確認する',
    '操作ログの実行者が、操作した本人になっている（2026/10/15 確認済み）',
    '見積書作成ですぐに新しいタブが開き、完了後に書類が表示される（2026/10/15 確認済み）',
    '一般社員・各承認者のアカウントで、サイドバーの書き込み操作が従来どおり動く',
    '承認者のメールアドレスを添えたリクエストを一般社員が送っても、一般社員として扱われ「権限がありません」になる']],
  ['開発', ['開発環境から取り込んだ機能を確認する（ユーザーテスト・チェックリストのG-13〜G-17・L区分）',
    '通知方法を社員DBの「メール通知」「Chat通知」チェックボックスで選べる',
    '差し戻し時に作成者へ通知が届く。最終承認完了の通知に承認者名が入る',
    '旧一覧からの案件移行（点検・取り込み・番号の引き継ぎ）']],
  ['開発', 'フェーズ0で編集権限の扱いを「スクリプトの切り離し」に決めた場合は、ここで規模を見積もり、本稼働日への影響を決裁者と相談する'],
]));
c.push(gap());
c.push(done('開発環境で9月17日チェックリストがすべて確認済み。ドキュメントが実装と一致している。'));

// フェーズ2
c.push(h1('フェーズ2　本番環境の構築'));
c.push(period('2026/10/26〜10/30'));
c.push(p('`初期設定マニュアル_v2.4.docx` の章番号に沿って進めます。**すべて管理用アカウントで作業します。**'));
c.push(table(['#', '作業', '担当', 'マニュアル'], [
  ['2-1', 'メインスプレッドシートを作成し、案件シート（`{n}期_表示`／`{n}期_全案件DB`／原本）・社員DB・取引先DB・config・操作ログを用意', '管理', '3.1〜3.3'],
  ['2-2', '社員DBに全社員とロール・通知方法（L・M列のチェックボックス）を登録', '管理', '3.4'],
  ['2-3', '取引先DBに本番の取引先を登録', '管理', '3章'],
  ['2-4', 'Driveフォルダを作成（見積書・請求書・納品書・アーカイブ・バックアップ）', '管理', '4章'],
  ['2-5', '本番用テンプレート3種を配置し、転記先セル（9月17日変更後の位置）とPDF出力範囲を整える', '管理', '5章'],
  ['2-6', '社印画像を配置し、共有リンクを取得', '管理', '6章'],
  ['2-7', '本番用の clasp プロジェクトへコードを反映（`clasp push --user admin -f`）', '開発', '7章'],
  ['2-8', 'Webアプリをデプロイし、デプロイIDとURLを控える', '開発・管理', '8章'],
  ['2-9', 'スクリプトプロパティを設定（READMEに記載の全キー）', '管理', '9章'],
  ['2-10', 'メニュー「WebアプリのURLを設定する」でURLを登録', '管理', '8章'],
  ['2-11', '`installTriggers` を**管理用アカウントで1回だけ**実行し、認可する', '管理', '10.1'],
  ['2-12', '「案件シートの構成を確認・修復する」を実行（列数が37列になっていること）', '管理', '3.5'],
  ['2-13', '案件シートの自動書き込み列を保護（AK列まで）', '管理', '3章'],
  ['2-14', '「取引先名の選択肢を更新する」を実行し、プルダウンを生成', '管理', '—'],
  ['2-15', '初期設定マニュアル11章「動作確認チェックリスト」を実施', '管理・開発', '11章'],
  ['2-16', 'メインスプレッドシート・Driveフォルダを利用者に共有（権限の範囲を確認）', '管理', '—'],
], [mm(15), W - mm(15) - mm(24) - mm(20), mm(24), mm(20)], { center: [0, 3] }));
c.push(gap());
c.push(tasks([
  ['開発', '本番用ディレクトリ（`../my-gas-project-prod/`）と `.clasp.json`、デプロイIDを手順書として残す'],
  ['管理', 'トリガー一覧に、管理用アカウントのトリガー（onEdit・日次0:30・朝10:00・サイドバー自動表示のonOpen）が1つずつだけあることを確認する'],
]));
c.push(gap());
c.push(done('初期設定マニュアル11章の動作確認がすべてOK。'));

// フェーズ3
c.push(h1('フェーズ3　ユーザーテスト'));
c.push(period('2026/11/2〜11/13'));
c.push(p('`ユーザーテスト_チェックリスト_v2.4.docx`（194項目）を本番環境で実施します。L区分「案件移行」（6項目）は、テストデータを消した後のフェーズ5で実施します。'));
c.push(h2('準備'));
c.push(tasks([
  ['管理', 'テスト参加者が各自1回「ユーザー用システム設定 > サイドバー自動表示を有効にする」を実行する'],
  ['管理', 'テスト用取引先を取引先DBに登録する（同じ社名で担当者が異なる組を1組含める）'],
  ['管理', 'テスト案件名の先頭に「テスト_」を付けるルールを周知する'],
  ['開発', '不具合報告の受け付け方法を決める（チェックリスト3章「不具合報告シート」に記入して共有）'],
]));
c.push(h2('実施の順番'));
c.push(table(['日程', '区分', '所要', '主な担当'], [
  ['1日目', 'A. 初期設定／B. 案件登録', '25分', '管理'],
  ['1〜2日目', 'C. 見積書', '50分', '一般社員・見積書承認者'],
  ['2日目', 'D. 請求書／E. 納品書／E2. 請求後の再見積もり', '55分〜', '一般社員・請求書承認者'],
  ['3日目', 'F. 案件のクローズ', '30分', '最終承認者'],
  ['3日目', 'G. 通知', '40分', '総務・全員'],
  ['4日目', 'H. 権限／I. 操作ログ', '40分', '一般社員・管理'],
  ['4日目', 'K. 異常系・境界値', '20分', '開発・管理'],
  ['翌日以降', 'J. データ保全（日次バックアップの生成を翌日に確認）', '20分', '管理'],
], [mm(22), W - mm(22) - mm(18) - mm(44), mm(18), mm(44)], { center: [2] }));
c.push(h2('定期レポートの確認日'));
c.push(p('定期レポートは送信日が決まっているため、テスト期間の日付に合わせて確認します。'));
c.push(table(['レポート', '送信タイミング', 'テスト期間中の確認'], [
  ['請求予定レポート', '毎月5日・20日 0:30', '11/5 に届く'],
  ['請求漏れリマインド', '毎月5日 10:00', '11/5 に届く（請求予定を過ぎた未請求のテスト案件を前日までに用意する）'],
  ['放置書類アラート', '毎日 10:00（作成中のまま7日経過した書類があるとき）', '11/2 に作成中の書類を残しておき、11/9 以降に届く'],
  ['未承認案件レポート', '毎月25日 10:00', ['期間外のため、スクリプトエディタから `runScheduledNotificationsForTest` を実行して確認する（チェックリストG-8・G-9）', '実際の送信日の動作は、本稼働後の12/25に確認する']],
], [mm(34), mm(46), W - mm(80)]));
c.push(gap(80));
c.push(note('未承認案件レポート・請求漏れリマインド・放置書類アラートは、`runScheduledNotificationsForTest` で送信日を待たずにまとめて送れます。上表は、実際の送信日に自動で届くことの確認です。'));
c.push(gap());
c.push(tasks([
  ['管理', '上表の日付に、各レポートが届くことを確認する'],
  ['管理', '日次0:30の処理（翌期準備・選択肢更新・アーカイブ・バックアップ）がエラーなく動いていることを、実行ログで確認する'],
  ['全員', 'テスト結果を4.1「実施結果サマリ」に集計する'],
]));
c.push(gap());
c.push(done('L区分を除く188項目すべてにOK／NGが記入されている（未実施なし）。NG項目が不具合報告シートに記録されている。'));

// フェーズ4
c.push(h1('フェーズ4　不具合修正・再テスト'));
c.push(period('2026/11/16〜11/20'));
c.push(tasks([
  ['開発', 'NG項目を「本稼働前に必須で直す」「稼働後でよい」「運用で回避」に仕分け、決裁者と合意する'],
  ['開発', '必須の不具合を修正し、開発環境で確認してから本番へ反映する（`clasp push` の後、`clasp deploy -i <デプロイID>` で**同じURLのまま再デプロイ**）'],
  ['利用・管理', '修正した項目と、その周辺の項目を再テストする'],
  ['開発', '修正内容をドキュメント（`docs/generators/`）に反映し、マニュアルを再生成する'],
  ['決裁', 'チェックリスト4.3「総合判定」で本番運用の可否を判定する'],
]));
c.push(gap());
c.push(done('総合判定が「可」または「条件付きで可」。条件付きの場合、条件と対応期限が記録されている。'));

// フェーズ5
c.push(h1('フェーズ5　本番切り替えの準備'));
c.push(period('2026/11/24〜11/27'));
c.push(tasks([
  ['管理', ['テストデータを消去する（スクリプトエディタから `runResetForUserTest` を実行）',
    '案件シートのデータ行・案件番号カウンター・操作ログ・当期フォルダ内の書類ファイルが消えることを確認',
    '表示シートの見出しが崩れた場合は `runRestoreUiHeaderRow2` を実行']],
  ['管理', 'テスト用取引先・テスト用社員を各DBから削除する（取引先DBはリセットの対象外のため手作業）'],
  ['管理', '「取引先名の選択肢を更新する」を再実行する'],
  ['管理', ['案Bの場合、旧一覧の案件を取り込む（初期設定マニュアル12.7、チェックリストL区分）',
    '「移行データ」シートに、フェーズ0でまとめた表を貼り付ける',
    '`runCheckMigrationSheet` で点検し、不備が無くなるまで直す（点検では何も書き込まれない）',
    '`runImportCasesFromMigrationSheet` で取り込み、次の新規案件が取り込んだ最大の番号の続きになることを確認する']],
  ['管理', '案Aの場合は、案件番号が `{期}-001` から始まることを確認する'],
  ['管理', '全利用者に、メインスプレッドシートとDriveフォルダの共有が行き渡っているか確認する'],
  ['開発・管理', ['利用者向け説明会を開く（ユーザー用マニュアルv2.4を配布）',
    'サイドバー自動表示の有効化（各自1回。自分で有効にしないと自動では開かない）',
    '通知はメール・Chatworkのどちらで届くか（社員DBの設定どおり）',
    '案件の登録と自動採番。取引先名は「社名（担当者名）」から選ぶこと',
    '見積書→請求書→納品書の流れ、承認依頼、差し戻し、再作成、請求後の再見積もり',
    '自動書き込み列や書類のロックされた範囲は編集できないこと']],
  ['管理', '問い合わせ窓口と、障害時の連絡先を周知する'],
]));
c.push(gap());
c.push(done('テストデータが残っておらず、全利用者が説明会を受けてサイドバー自動表示を有効にしている。'));

// フェーズ6
c.push(h1('フェーズ6　本稼働・初期サポート'));
c.push(period('2026/12/1〜12月末'));
c.push(h2('稼働初週（毎日）'));
c.push(tasks([
  ['管理', '操作ログにエラーが記録されていないか確認する'],
  ['管理', 'スクリプトエディタの「実行数」で、日次・朝の定期処理が失敗していないか確認する'],
  ['管理', 'バックアップファイルが毎日作られているか確認する'],
  ['開発', '利用者からの問い合わせ・不具合に対応する'],
]));
c.push(h2('稼働1か月目（週1回）'));
c.push(tasks([
  ['管理', '上の確認を週1回に減らして続ける'],
  ['管理', '12/5（請求予定レポート・請求漏れリマインド）、12/20（翌月分の請求予定レポート）、12/25（未承認案件レポート）の内容が実態と合っているか確認する'],
  ['開発', '問い合わせの多い操作を洗い出し、マニュアルを補足する'],
  ['決裁', '12月末に稼働1か月の振り返りを行い、改善要望を整理する'],
]));
c.push(gap());
c.push(done('1か月分の見積→請求→納品→最終承認が、大きな不具合なく回っている。'));

// フェーズ7
c.push(h1('フェーズ7　年間の節目の確認'));
c.push(p('一度しか起きない処理のため、日付を予定表に入れておきます。'));
c.push(table(['時期', '確認すること', '担当'], [
  ['2027/4/1 の翌朝', '翌期の `{n}期_表示`／`{n}期_全案件DB` と各Driveフォルダが自動作成されている', '管理'],
  ['2027/5/1', '新規案件が新しい期の番号（`{n}-001`）で採番される。前期の進行中案件も引き続き操作できる', '管理'],
  ['前期の全案件が完了した翌日', '前期のシートが `案件データ_{n}期` へ退避され、本体から消えている', '管理'],
  ['随時', 'コードを変えたら、`clasp push` と**Webアプリの再デプロイ**を必ずセットで行う', '開発'],
], [mm(40), W - mm(40) - mm(18), mm(18)], { center: [2] }));

// リスク
c.push(h1('リスク・懸念事項'));
c.push(h2('1. Webアプリ入口の本人確認'));
c.push(callout('', ['**2026/10/15 に対策済み。**開発環境で、実行者が本人として記録されることを確認済みです。']));
c.push(gap(80));
c.push(p('Webアプリは `access: ANYONE`（Googleにログインしていれば誰でも）でデプロイしています。これまでの `doPost`（WebAppEntry.js）は、リクエストに書かれたメールアドレスをそのまま実行者として扱っていたため、URLを知っている人が承認者のアドレスを名乗れば、承認者として操作できました。'));
c.push(p('社員DBを保護していても、この方法は防げませんでした。社員DBを書き換えるのではなく、社員DBに正しく登録されている承認者のアドレスを名乗る方法だからです。'));
c.push(p('現在は、呼び出し元のアクセストークン（Googleが本人にだけ発行するログイン情報）をGoogleに照会し、その持ち主のメールアドレスを実行者として使います（`verifyCallerToken_`）。'));
c.push(table(['', '対策前', '対策後'], [
  ['実行者の決め方', 'リクエストに書かれたメールアドレス', 'アクセストークンの持ち主（Googleに照会）'],
  ['他人へのなりすまし', 'できる', 'できない'],
  ['URLを知っている人ができること', '誰の権限でも操作できる', '自分の権限の範囲でのみ操作できる（サイドバーと同じ）'],
  ['処理時間', '—', '照会のため通信が1回増える（同じ人の照会結果は10分間覚えておく）'],
], [mm(44), (W - mm(44)) / 2, (W - mm(44)) / 2]));
c.push(gap(80));
c.push(p('あわせて、サイドバーの自動表示についても同日に直しました。自動表示のトリガーは誰がスプレッドシートを開いても登録した人の権限で動くため、管理用アカウントのトリガーで開いたサイドバーからの操作が、管理用アカウントとして実行されていました。現在は、トリガーを登録した本人が開いたときだけサイドバーを表示します。'));
c.push(h2('2. 編集者はスクリプトも編集できる'));
c.push(p('メインスプレッドシートに付属したスクリプトは、スプレッドシートの編集権限があれば「拡張機能 > Apps Script」から開いて書き換えられます。社員は案件を入力するため、全員が編集権限を持っています。'));
c.push(p('管理用アカウントで登録したトリガー（onEdit・日次・朝）は、**保存されている最新のコードを管理用アカウントの権限で実行します。**そのため、悪意を持ってコードを書き換えれば、管理用アカウントとして任意の処理を動かせます。社員DBの保護や、1の本人確認では防げません。'));
c.push(table(['対策の候補', '効果と注意点'], [
  ['編集権限を持つ人を信頼できる範囲に限ったうえで、`clasp pull` で本番のコードを定期的に取得し、リポジトリとの差分がないか確認する', '運用での対処。コード変更は不要。書き換えを防ぐのではなく、後から気づく仕組み'],
  ['スクリプトをスプレッドシートから切り離し、編集権限を管理用アカウントと開発担当だけにする（単独のスクリプトとして作り、アドオンとして読み込む）', '根本的な対策。メニュー・サイドバー・トリガーの作り直しが必要で、規模が大きい。本稼働日に影響する可能性がある'],
], [W / 2, W / 2]));
c.push(h2('3. Workspace未契約による制約'));
c.push(bullet('シートを直接編集した人は特定できない（操作ログには「実行者不明」と記録される）'));
c.push(bullet('書類のオーナーは管理用アカウントに固定される。管理用アカウントを失うと全書類に影響するため、**認証情報の保管と2段階認証の復旧手段を必ず用意する**'));
c.push(h2('4. トリガーの重複'));
c.push(p('`installTriggers` を管理用アカウント以外が実行すると、1回の編集で処理が人数分走ります。スクリプトからは検知できないため、実行するのは管理用アカウントのみと周知します。'));
c.push(h2('5. 再デプロイ忘れ'));
c.push(p('`clasp push` だけでは、書き込み系の操作に変更が反映されません。GitHub Actions（`.github/workflows/deploy.yml`）が行うのは**開発環境（`.clasp.json` の `1cHuC…`）への push のみ**で、Webアプリの再デプロイと本番への反映は手作業です。'));
c.push(p('また、開発環境のコードがリポジトリより進んでいたことがありました（2026/10/15 に取り込み済み）。開発環境で直接コードを変えた場合は、必ずリポジトリにも反映します。'));

// 関連ドキュメント
c.push(h1('関連ドキュメント'));
c.push(table(['ファイル（docs/ 配下）', '使うフェーズ'], [
  ['9月17日_修正点と動作確認チェックリスト.txt', '1'],
  ['9月4日_修正点と動作確認チェックリスト.txt', '1'],
  ['初期設定マニュアル_v2.4.docx', '2・7'],
  ['ユーザーテスト_チェックリスト_v2.4.docx', '3・4'],
  ['ユーザー用マニュアル_v2.4.docx', '5'],
  ['仕様書_v2.4.docx', '全体'],
  ['README.md（リポジトリ直下）', '2・開発'],
], [W - mm(30), mm(30)], { center: [1] }));

// 改訂履歴
c.push(h1('改訂履歴'));
c.push(table(['日付', '内容', '担当'], [
  ['2026/10/13', '初版', '【要確認】'],
  ['2026/10/15', '定期レポートの確認日を追記し、社内標準の様式でWord化', '【要確認】'],
  ['2026/10/15', 'リスク1を対策済みに更新し、リスク2（スクリプトの編集権限）を追加', '【要確認】'],
  ['2026/10/15', '開発環境から取り込んだ機能（通知方法の選択・案件移行など）とサイドバーの修正を反映。案Bを番号の引き継ぎに改め、定期レポートの確認方法を訂正', '【要確認】'],
], [mm(28), W - mm(28) - mm(30), mm(30)]));

// ---------- 文書 ----------

function header() {
  const logo = fs.readFileSync(path.join(__dirname, 'assets', 'logo-color.png'));
  return new Header({
    children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: W }],
      children: [
        new ImageRun({ type: 'png', data: logo, transformation: { width: 104, height: 19 } }),
        new TextRun({ text: `\t${DOC_TITLE}`, size: 16, color: COLOR.note }),
      ],
    })],
  });
}

function footer() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: COLOR.note }),
        new TextRun({ text: ' / ', size: 16, color: COLOR.note }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: COLOR.note }),
      ],
    })],
  });
}

module.exports = () => new Document({
  title: DOC_TITLE,
  styles: {
    default: {
      document: {
        run: { font: { ascii: FONT, eastAsia: FONT, hAnsi: FONT }, size: 21, color: COLOR.text },
        paragraph: { spacing: { line: 360 } },
      },
      title: {
        run: { font: FONT, size: 44, bold: true, color: COLOR.main },
        paragraph: { spacing: { after: 80 } },
      },
      heading1: {
        run: { font: FONT, size: 32, bold: true, color: COLOR.main },
        paragraph: { spacing: { before: 480, after: 160 }, keepNext: true },
      },
      heading2: {
        run: { font: FONT, size: 26, bold: true, color: COLOR.main },
        paragraph: { spacing: { before: 320, after: 120 }, keepNext: true },
      },
    },
  },
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '・', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 240 } } } }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: mm(210), height: mm(297) },
        margin: { top: mm(30), bottom: mm(25), left: mm(25), right: mm(25), header: mm(12), footer: mm(12) },
      },
    },
    headers: { default: header() },
    footers: { default: footer() },
    children: c,
  }],
});
