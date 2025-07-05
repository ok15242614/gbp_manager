# GBP Manager

Google スプレッドシートの口コミデータを月別にGoogleドキュメントとして出力するGoogle Apps Script集です。

## 主な機能
- 指定年月の口コミのみ抽出し、Googleドライブの「yyyy年M月」サブフォルダに保存
- 店舗ごと・全店舗まとめのドキュメント出力（切替可）
- コメント原文抽出・日付フォーマット変換もサポート

## ファイル構成
- `docGenerator.js` : 口コミデータの月別抽出・出力
- `extractComments.js` : コメント原文抽出
- `formatDate.js` : 日付フォーマット変換
- `appsscript.json` : GASプロジェクト設定

## 使い方
1. Apps Scriptに本リポジトリのスクリプトを貼り付け
2. スクリプトプロパティで `SPREADSHEET_ID` と `FOLDER_ID` を設定
3. `generateDoc()` で直前の月、`generateDoc('2024年6月', false)` で2024年6月分まとめのみ出力

## 注意
- 実行前にバックアップ推奨
- スプレッドシートのタイムゾーンは「(GMT+09:00) 東京」に設定 