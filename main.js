function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('output doc')
    .addItem('口コミデータをGoogleドキュメントに出力', 'generateDoc')
    .addToUi();
}