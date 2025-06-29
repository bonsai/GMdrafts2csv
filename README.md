![image](https://github.com/user-attachments/assets/a080e1c5-0572-4284-9591-8ad2b5d4196e)

# Gmail下書きスプレッドシート連携 GAS

Gmail の下書きメッセージを取得して Google スプレッドシートに保存する Google Apps Script プロジェクトです。

## 機能

- Gmail の下書きメッセージを自動取得
- 下書きの詳細情報（件名、宛先、本文など）をスプレッドシートに保存
- エラーハンドリングとログ出力
- 日本語対応

## セットアップ

### 1. Google Apps Script プロジェクトの作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 「新しいプロジェクト」をクリック
3. プロジェクト名を設定（例：「Gmail下書き管理」）

### 2. コードのデプロイ

1. `Code.gs` の内容をメインの Code.gs ファイルにコピー
2. `appsscript.json` の内容をマニフェストファイルに設定
   - エディター画面左側の「プロジェクトの設定」→「appsscript.json マニフェスト ファイルをエディタで表示する」をチェック

### 3. スプレッドシートの準備

1. 新しい Google スプレッドシートを作成
2. スプレッドシートの URL からスプレッドシート ID を取得
3. `Code.gs` 内の `SPREADSHEET_ID` 変数に実際の ID を設定

### 4. 権限の許可

1. スクリプトエディターで `main` 関数を選択
2. 「実行」ボタンをクリック
3. 権限の許可を求められた場合は、Gmail と Google Sheets のアクセス許可を与える

## 使用方法

### 基本的な使用方法

1. `main()` 関数を実行すると、Gmail の下書きを取得してスプレッドシートに保存されます
2. スプレッドシート内に「Gmail下書き」という名前のシートが自動作成されます

### 定期実行の設定

1. スクリプトエディター左側の「トリガー」をクリック
2. 「トリガーを追加」を選択
3. 実行する関数：`main`
4. イベントソース：時間主導型
5. お好みの間隔を設定（例：1時間ごと、毎日など）

## 出力されるデータ

スプレッドシートには以下の情報が保存されます：

| 列名 | 説明 |
|------|------|
| 下書きID | Gmail下書きの一意識別子 |
| 件名 | メールの件名 |
| 宛先 | To フィールド |
| CC | CC フィールド |
| BCC | BCC フィールド |
| 作成日時 | 下書きの作成日時 |
| 添付ファイル数 | 添付されているファイルの数 |
| 本文プレビュー | 本文の最初の100文字 |
| 取得日時 | データを取得した日時 |

## 関数の説明

### `main()`
メイン実行関数。Gmail下書きの取得からスプレッドシートへの保存まで一連の処理を実行します。

### `getDrafts()`
Gmail の下書きを取得し、必要な情報を抽出します。

### `saveToSheet(spreadsheetId, drafts)`
取得した下書きデータをスプレッドシートに保存します。

### `checkDraftCount()`
現在の下書き数を確認するテスト関数です。

## 注意事項

- Gmail API の制限により、一度に大量の下書きを処理する場合は実行時間制限に注意してください
- スプレッドシートIDは実際のものに変更してください
- 権限の許可が必要な場合があります

## トラブルシューティング

### よくある問題

1. **「権限が不足しています」エラー**
   - スクリプトの初回実行時に権限許可を行ってください

2. **「スプレッドシートが見つかりません」エラー**
   - `SPREADSHEET_ID` が正しく設定されているか確認してください

3. **「実行時間の上限を超えました」エラー**
   - 下書きの数が多い場合は、バッチ処理を検討してください

## ライセンス

このプロジェクトは MIT ライセンスの下で公開されています。
