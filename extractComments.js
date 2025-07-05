/**
 * 口コミの原文を抽出する（全シート対応版）
 * スプレッドシート内の全てのシートに対して処理を実行
 * Google翻訳されたレビューから元の言語のテキストを抽出する
 * - (Original) があれば、その後の文章を抽出する。
 * - (Original) がなく、(Translated by Google) があれば、その前の文章を抽出する。
 */
function extractComments() {
  // --- 設定項目 ---
  // 口コミが入力されている列の番号（A列=1, B列=2, C列=3, D列=4...）
  const reviewColumn = 4;   // ★口コミが入力されている列番号に変更してください (A列=1, B列=2, ...)
  
  // データ処理を開始する行番号（通常はヘッダー行の次の行）
  const startRow = 2;       // ★処理を開始する行番号に変更してください
  
  // 処理から除外するシート名の配列（必要に応じて追加）
  const excludeSheets = ['設定', 'テンプレート', 'マスタ']; // ★除外したいシート名があれば追加してください
  // --- 設定はここまで ---

  // アクティブなスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // スプレッドシート内の全てのシートを取得
  const allSheets = spreadsheet.getSheets();
  
  // 処理結果を記録するための変数
  let processedSheets = 0;  // 処理されたシート数
  let totalProcessedRows = 0; // 処理された行数の合計
  let skippedSheets = [];   // スキップされたシート名の配列
  
  console.log(`処理開始: 全${allSheets.length}シートを対象に処理します`);
  
  // 各シートに対して処理を実行
  allSheets.forEach((sheet, index) => {
    const sheetName = sheet.getName();
    
    console.log(`[${index + 1}/${allSheets.length}] シート "${sheetName}" を処理中...`);
    
    // 除外対象のシートかチェック
    if (excludeSheets.includes(sheetName)) {
      console.log(`  → スキップ: "${sheetName}" は除外対象です`);
      skippedSheets.push(sheetName);
      return; // このシートの処理をスキップ
    }
    
    try {
      // 個別シートの処理を実行
      const processedRows = processSheet(sheet, reviewColumn, startRow);
      
      if (processedRows > 0) {
        processedSheets++;
        totalProcessedRows += processedRows;
        console.log(`  → 完了: ${processedRows}行を処理しました`);
      } else {
        console.log(`  → スキップ: 処理対象データがありません`);
        skippedSheets.push(sheetName);
      }
      
    } catch (error) {
      // エラーが発生した場合の処理
      console.error(`  → エラー: シート "${sheetName}" の処理中にエラーが発生しました`);
      console.error(`    エラー内容: ${error.message}`);
      skippedSheets.push(sheetName + ' (エラー)');
    }
  });
  
  // 処理結果をユーザーに通知
  let resultMessage = `処理完了！\n\n`;
  resultMessage += `処理されたシート数: ${processedSheets}/${allSheets.length}\n`;
  resultMessage += `処理された総行数: ${totalProcessedRows}行\n`;
  
  if (skippedSheets.length > 0) {
    resultMessage += `\nスキップされたシート:\n${skippedSheets.join('\n')}`;
  }
  
  console.log(resultMessage);
}

/**
 * 個別のシートに対して口コミ原文抽出処理を実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 処理対象のシート
 * @param {number} reviewColumn - 口コミが入力されている列番号
 * @param {number} startRow - 処理開始行番号
 * @return {number} 処理された行数
 */
function processSheet(sheet, reviewColumn, startRow) {
  // シートの最後の行を取得（データがある最後の行番号）
  const lastRow = sheet.getLastRow();
  
  // 処理対象のデータが存在しない場合は0を返す
  if (lastRow < startRow) {
    return 0; // 処理対象データがない場合
  }
  
  // 指定された列が存在するかチェック
  const lastColumn = sheet.getLastColumn();
  if (reviewColumn > lastColumn) {
    console.log(`    警告: 列${reviewColumn}が存在しません（最大列: ${lastColumn}）`);
    return 0;
  }

  // 処理対象の範囲を取得
  // startRow行目からlastRow行目まで、reviewColumn列の1列分のデータを取得
  const range = sheet.getRange(startRow, reviewColumn, lastRow - startRow + 1, 1);
  
  // 範囲内の全ての値を2次元配列として取得
  const values = range.getValues();

  // 各行のデータを処理して新しい値の配列を作成
  const newValues = values.map(row => {
    // row[0]は各行の最初の列（今回は口コミ列）の値
    let text = row[0];

    // 値が文字列でない、または空文字列の場合はそのまま返す
    if (typeof text !== 'string' || text === '') {
      return [text]; // 元の値をそのまま配列で返す
    }

    // Google翻訳で使用されるマーカー文字列を定義
    const originalMarker = '(Original)';      // 原文を示すマーカー
    const translatedMarker = '(Translated by Google)'; // 翻訳文を示すマーカー
    
    // 最優先: (Original) マーカーが含まれているかチェック
    const originalIndex = text.indexOf(originalMarker);
    if (originalIndex !== -1) {
      // (Original) マーカーが見つかった場合
      // マーカー位置 + マーカーの文字数分を足して、その後のテキストを抽出
      text = text.substring(originalIndex + originalMarker.length);
      return [text.trim()]; // 前後の空白を削除して配列で返す
    }
    
    // (Original) がなく、(Translated by Google) マーカーが含まれている場合
    const translatedIndex = text.indexOf(translatedMarker);
    if (translatedIndex !== -1) {
      // (Translated by Google) マーカーが見つかった場合
      // マーカーの前にあるテキスト（原文）を抽出
      text = text.substring(0, translatedIndex);
      return [text.trim()]; // 前後の空白を削除して配列で返す
    }
    
    // どちらのマーカーも見つからない場合
    // 既に原文のみとみなし、前後の空白のみを削除してそのまま返す
    return [text.trim()];
  });

  // 処理した新しい値をスプレッドシートの同じ範囲に書き込み
  // 元のデータを上書きする
  range.setValues(newValues);
  
  // 処理された行数を返す
  return lastRow - startRow + 1;
}