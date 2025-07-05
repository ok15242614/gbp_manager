function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GBP Manager')
    .addItem('口コミデータをGoogleドキュメントに出力', 'generateDoc')
    .addItem('コメント原文抽出', 'extractComments')
    .addItem('日付をJSTに変換', 'formatDate')
    .addToUi();
}