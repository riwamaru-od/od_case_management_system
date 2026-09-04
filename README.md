# 案件管理・見積書自動作成システム

OVER D-LIVE の案件管理と、見積書・請求書・納品書の作成から承認・PDF出力までを
Googleスプレッドシート上で完結させる Google Apps Script プロジェクトです。

案件シートの行を選択するとサイドバーが開き、そこから書類の作成・承認依頼・承認・
PDF出力までを一連の流れで操作します。書類はテンプレートから複製され、取引先DB・
社員DBの情報が自動で転記されます。

## 主な機能

| 機能 | 概要 |
| --- | --- |
| 案件の自動採番 | 取引先名・案件名・終了予定・請求予定・担当の5項目が揃った時点で `{期}-{連番}` を採番 |
| 書類の自動作成 | テンプレートから複製し、取引先DB・社員DBの情報を転記。見積書→請求書→納品書と内容を引き継ぐ |
| 承認フロー | 作成 → 完成（承認依頼）→ 承認／差し戻し → PDF出力。見積書は承認依頼先を指定可能 |
| 社印の自動押印 | 見積書・請求書は承認時、納品書は作成時に押印。位置は行の高さから自動計算 |
| PDF出力 | 出力範囲を制御し、操作履歴欄を除いた本文のみを出力。出力者と日時を記録 |
| 見積・請求の差分検知 | 請求書が見積書と食い違うセルを、条件付き書式でリアルタイムに黄色表示 |
| 通知 | 承認依頼・承認完了を通知（総務はメール、それ以外はChatwork） |
| 定期レポート | 請求予定レポート、未承認案件レポート、請求漏れリマインド、放置書類アラート |
| 期の管理 | 毎年5月1日に期が切り替わり、1ヶ月前に翌期のシート・フォルダを自動作成 |
| アーカイブ・バックアップ | 過去期のシートを別ファイルへ退避。DBを日次でバックアップ |

## アーキテクチャ

### サイドバーの書き込み操作は、すべて管理用アカウントの権限で実行される

このプロジェクトの中核となる設計です。

サイドバーからの `google.script.run` 呼び出しは、常に**操作した本人の権限**で実行されます。
これは以下2つの問題を引き起こしていました。

1. 案件シートの自動書き込み列（ステータス・書類リンク等）を手入力から守るために列を保護すると、
   スクリプトからの書き込みまでブロックされてしまう
2. 作成された書類ファイルのオーナーが、ボタンを押した本人になってしまう

そこで、同じスクリプトを **Webアプリとして `executeAs: USER_DEPLOYING` でデプロイ**し、
書き込み系の処理はサーバー間通信（`UrlFetchApp`）でそのWebアプリを呼び出す構成にしています。
実処理は常に管理用アカウントの権限で実行されます。

```
サイドバー (JavaScript.html)
  │  google.script.run          ← 操作した本人の権限で実行
  ▼
SidebarController.js  api_*()
  │  callAsAdmin_()             ← AdminProxyService.js
  │  UrlFetchApp で自分自身のWebアプリへPOST
  ▼
WebAppEntry.js  doPost()        ← 管理用アカウントの権限で実行
  │
  ▼
各 *Service.js（実処理）
```

### 実行者の特定について

Webアプリを `access: ANYONE` でデプロイしているため、`doPost` の中では
`Session.getActiveUser()` が呼び出し元を返しません。そのため `callAsAdmin_` が
呼び出し元自身の実行コンテキストで取得したメールアドレスを `callerEmail` として転送し、
転送先で `ACTIVE_USER_EMAIL_OVERRIDE_`（Utils.js）に設定して本人として扱っています。

`access: DOMAIN` であればこの回避は不要ですが、後述のとおり実Google Workspaceドメインが
無いため利用できません。

## この環境固有の制約

実際の Google Workspace を契約していないため、以下が使えません。実装上の回避策も併記します。

| 使えないもの | 影響 | 回避策 |
| --- | --- | --- |
| `DriveApp.File.setOwner()` | 書類のオーナーを別アカウントへ移せない | Webアプリ経由で管理用アカウントとして作成する |
| Webアプリの `access: DOMAIN` | ドメイン所属を解決できず404になる | `access: ANYONE` にし、実行者は `callerEmail` で転送 |
| `e.user` / `Session.getActiveUser()`（onEditトリガー内） | シートを直接編集した人を特定できない | 操作ログに「シート直接編集（実行者不明）」と記録 |

## ファイル構成

### 入り口

| ファイル | 役割 |
| --- | --- |
| `Triggers.js` | メニュー、onOpen/onEdit、日次・朝の定期処理の入り口 |
| `SidebarController.js` | サイドバーから呼ばれる `api_*` とボタンの活性判定 |
| `WebAppEntry.js` | Webアプリの `doPost`。アクション名から実処理へ振り分ける |
| `AdminProxyService.js` | `callAsAdmin_`。Webアプリを呼び出すプロキシ |
| `Sidebar.html` / `JavaScript.html` / `Stylesheet.html` | サイドバーのUI |

### 業務ロジック

| ファイル | 役割 |
| --- | --- |
| `CaseService.js` | 案件行の採番・DB同期・中止・変更記録・番号順の整列 |
| `CaseRestoreService.js` | 案件中止・最終承認の取り消し |
| `ApprovalService.js` | 見積書・請求書の作成／完成／承認／差し戻し／再作成／PDF出力 |
| `DeliveryService.js` | 納品書の作成（承認フロー無し） |
| `FinalApprovalService.js` | 最終承認 |
| `DocumentService.js` | 書類ファイルの複製・フォルダ配置・PDF書き出し |
| `TemplateFillService.js` | テンプレートへの転記・社印・シート保護 |
| `QuoteInvoiceDiffService.js` | 見積書と請求書の差分ハイライト |
| `NotificationService.js` | メール／Chatwork通知の組み立てと送信 |
| `ReportService.js` | 定期レポート・リマインド・アラート |

### 基盤・設定

| ファイル | 役割 |
| --- | --- |
| `Constants.js` | 列定義・ロール・ステータス・テンプレートのセル番地など |
| `Config.js` | スクリプトプロパティの読み出し、Drive上のフォルダ・ファイル取得 |
| `DocTypes.js` | 見積書・請求書・納品書の差分を吸収する設定オブジェクト |
| `Utils.js` | エラー、排他制御（`withLock_`）、共通ユーティリティ |
| `PeriodService.js` | 期の判定と、翌期のシート・フォルダの準備 |
| `RoleService.js` / `ClientService.js` | 社員DB・取引先DBの読み出し |
| `LogService.js` | 操作ログへの追記 |
| `ScheduleOptionsService.js` | 終了予定・請求予定のプルダウン生成と今月請求のハイライト |
| `SetupService.js` | 案件シートの構成の確認・修復、WebアプリURLの設定 |
| `ArchiveService.js` / `BackupService.js` | 過去期のアーカイブとDBのバックアップ |
| `TestDataResetService.js` | テストデータのリセット（管理者がエディタから手動実行） |

## セットアップ

詳細な手順は `docs/初期設定マニュアル_v2.2.docx` を参照してください。概要は以下のとおりです。

### 1. スクリプトプロパティの設定

「ファイル > プロジェクトの設定 > スクリプト プロパティ」で設定します。

| キー | 内容 |
| --- | --- |
| `MAIN_SPREADSHEET_ID` | メインスプレッドシートのID |
| `SHEET_NAME_CLIENT_DB` / `SHEET_NAME_STAFF_DB` | 取引先DB・社員DBのシート名 |
| `FOLDER_ID_QUOTE_ROOT` / `FOLDER_ID_INVOICE_ROOT` / `FOLDER_ID_DELIVERY_ROOT` | 各書類のルートフォルダID |
| `FOLDER_ID_ARCHIVE_ROOT` / `FOLDER_ID_BACKUP_ROOT` | アーカイブ・バックアップの保存先 |
| `QUOTE_TEMPLATE_FILE_ID` / `INVOICE_TEMPLATE_FILE_ID` / `DELIVERY_TEMPLATE_FILE_ID` | 各テンプレートのファイルID |
| `COMPANY_SEAL_IMAGE_URL` | 社印画像の共有リンク |
| `CHATWORK_API_TOKEN` | Chatwork APIトークン |
| `ADMIN_TRIGGER_ACCOUNT_EMAIL` | 管理用アカウントのメールアドレス |
| `WEBAPP_URL` | WebアプリのURL（`.../exec`）。メニューからも設定可 |

### 2. Webアプリのデプロイ

`appsscript.json` の `webapp` 設定（`access: ANYONE` / `executeAs: USER_DEPLOYING`）のまま、
**管理用アカウントで**デプロイします。発行されたURLを、メニュー
「管理者用メニュー > WebアプリのURLを設定する」で登録します。

> **コードを変更したら再デプロイが必要です。** 書き込み操作はWebアプリ経由で実行されるため、
> `clasp push` だけでは反映されません。

### 3. トリガーの登録

管理用アカウントでスクリプトエディタから `installTriggers` を1回実行します。

| トリガー | 内容 |
| --- | --- |
| `onEditInstallable` | 自動採番、変更記録、番号順の整列 |
| `dailyScheduledTasks`（0:30） | 翌期の準備、選択肢更新、アーカイブ、バックアップ、請求予定レポート |
| `morningScheduledTasks`（10:00） | 未承認案件レポート、請求漏れリマインド、放置書類アラート |

> `onEditInstallable` と時間主導型トリガーは**必ず1人だけ**が登録してください。
> 複数人が登録すると1回の編集で人数分の処理が走ります。Apps Scriptの仕様上、
> 他人が登録したトリガーはスクリプトから見えないため、重複はコード側で検知できません。

利用者は各自1回、メニュー「ユーザー用システム設定 > サイドバー自動表示を有効にする」を
実行します。

### 4. 案件シートの構成確認

メニュー「管理者用メニュー > 案件シートの構成を確認・修復する」を実行し、
列数と見出しを揃えます。列を追加する変更のあと、必ず実行してください。

## 開発

[clasp](https://github.com/google/clasp) を使い、開発用と本番用の2つのApps Scriptプロジェクトへ
それぞれ反映します。本番は専用の管理用アカウントで運用するため、clasp の名前付き認証情報
（`--user`）で使い分けています。

```bash
# 開発環境へ反映
clasp push -f

# 本番環境へ反映（別ディレクトリに本番用の .clasp.json を置いている）
cp <変更したファイル> ../my-gas-project-prod/
cd ../my-gas-project-prod
clasp push --user admin -f

# 本番のWebアプリを再デプロイ（同じURLのまま新しいバージョンにする）
clasp deploy --user admin -i <デプロイID> -d "変更内容"
```

`.claspignore` により、`clasp push` の対象はトップレベルの `*.js` / `*.html` /
`appsscript.json` のみです。`docs/` は対象外です。

### 設計上の約束ごと

- 書き込みを伴う処理は `withLock_()` で囲む。処理途中で失敗しても中途半端な状態が
  残らないよう、検証をすべて済ませてから書き込む
- ユーザー向けのエラーは `AppError_(code, message)` で投げる。想定外の例外は
  `withLock_` が捕捉して操作ログへ記録する
- サイドバーのボタンの活性制御はUX目的。権限の判定は必ずサーバー側（`assertRole_`）でも行う
- 列の追加・セル番地の変更は `Constants.js` の修正だけで完結させる

## ドキュメント

`docs/` に運用ドキュメントを置いています。

| ファイル | 内容 |
| --- | --- |
| `仕様書_v2.2.docx` | 機能仕様 |
| `初期設定マニュアル_v2.2.docx` | 導入手順 |
| `ユーザー用マニュアル_v2.2.docx` | 利用者向けの操作手順 |
| `ユーザーテスト_チェックリスト_v2.1.docx` | ユーザーテスト用のチェックリスト |
| `9月4日_修正点と動作確認チェックリスト.txt` | 直近の修正内容と動作確認手順 |
