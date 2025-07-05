/**
 * 全シートのA列の日付をUTCからJSTに変換し、yyyy/MM/dd形式で表示する
 * 
 * 前提条件：
 * - スプレッドシートのタイムゾーン設定が「(GMT+09:00) 東京」である必要があります
 * - [ファイル] > [設定] > [タイムゾーン] から設定を確認してください
 */
function formatDate() {
  // 定数定義
  const TARGET_TIMEZONE = 'JST';
  const DISPLAY_FORMAT = 'yyyy/MM/dd';
  const COLUMN_RANGE = 'A:A';
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = spreadsheet.getSheets();
    
    // 各シートのA列を処理
    allSheets.forEach(sheet => {
      const columnRange = sheet.getRange(COLUMN_RANGE);
      const originalValues = columnRange.getValues();
      
      // 各行の値を変換
      const convertedValues = originalValues.map(row => {
        const cellValue = row[0];
        
        // 空のセルまたは日付以外の値はそのまま返す
        if (cellValue === '' || cellValue == null || 
            !(cellValue instanceof Date) || isNaN(cellValue.getTime())) {
          return [cellValue];
        }
        
        // UTC日付をJST文字列に変換
        const jstString = Utilities.formatDate(
          cellValue, 
          TARGET_TIMEZONE, 
          'yyyy/MM/dd HH:mm:ss'
        );
        
        // JST文字列から新しい日付オブジェクトを作成
        // （スプレッドシートのタイムゾーン設定がJSTの場合、正しく解釈される）
        const jstDateObject = new Date(jstString);
        
        return [jstDateObject];
      });
      
      // 変換結果をシートに反映し、表示形式を設定
      columnRange
        .setValues(convertedValues)
        .setNumberFormat(DISPLAY_FORMAT);
    });
    
  } catch (error) {
    console.error(`日付フォーマット処理中にエラーが発生しました: ${error.toString()}`);
    throw error; // エラーを再スローして呼び出し元に通知
  }
}