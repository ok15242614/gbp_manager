function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ドキュメント出力')
    .addItem('ドキュメント出力', 'generateDoc')
    .addToUi();
}