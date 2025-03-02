# 🚀 2025FC GAS Checkpoint Scoring

## 📌 概要
このリポジトリは、Google Apps Script (GAS) を使用して、オリエンテーリングのチェックポイントのスコアを集計するシステムを管理するためのものです。

## 🎯 機能
✅ 各チェックポイントのスコアを集計
✅ 班ごとの合計得点を計算
✅ Google スプレッドシートと連携してデータを整理
✅ 実行履歴の記録

## 🛠 使用方法
1. Google Apps Script のエディタで新しいプロジェクトを作成
2. `collectCheckpointResults.js` のコードをコピーして貼り付け
3. **スプレッドシートのIDを設定**（`YOUR_SPREADSHEET_ID_HERE` を適切な値に変更）
   - GASコード内の `SpreadsheetApp.openById("...")` の部分にスプレッドシートのIDを入力してください。
   - 各チェックポイントのスプレッドシートIDも適宜変更してください。
4. 必要に応じて `summarizeTotalScores.js` も追加
5. スクリプトを実行してスコアを集計

## 📂 ファイル構成
- `collectCheckpointResults.js` - 各チェックポイントのスコアを集計
- `summarizeTotalScores.js` - 各班の最終合計スコアを計算

## 📝 注意事項
- Google Apps Script の「スクリプトプロパティ」機能を活用すると、スプレッドシートIDを直接コードに記載せずに管理できます。
- **スプレッドシートのIDを公開リポジトリに含めないよう注意してください。**
- スプレッドシートの権限設定を適切に行い、スクリプトがデータを読み書きできるようにしてください。

## 📢 変更履歴
- **v1.0** - 初回リリース（基本的なスコア集計機能を実装）

## 🤝 貢献方法
1. フォークしてブランチを作成 (`git checkout -b feature-branch`)
2. 変更をコミット (`git commit -m 'Add new feature'`)
3. プッシュ (`git push origin feature-branch`)
4. プルリクエストを作成

## 📧 問い合わせ
ご質問やフィードバックがありましたら、Issue を作成するか、リポジトリ管理者までご連絡ください！

