function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ドキュメント出力')
    .addItem('ドキュメント出力', 'generateDoc')
    .addItem("出力先フォルダ指定", "selectOutputFolder")
    .addToUi();
}