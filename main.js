function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ドキュメント出力')
    .addItem('ドキュメント出力', 'generateDoc')
    .addItem('年・月を指定して出力', 'selectOutputMonth')
    .addItem("出力先のフォルダを設定", "selectOutputFolder")
    .addToUi();
}