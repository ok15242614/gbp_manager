function generateDoc() {
    // スクリプトプロパティからIDを取得
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  
    // IDが設定されているか確認
    if (!spreadsheetId) {
      Logger.log('エラー: スプレッドシートIDがスクリプトプロパティに設定されていません。');
      SpreadsheetApp.getUi().alert('エラー', 'スクリプトプロパティにSPREADSHEET_IDを設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    if (!folderId) {
      Logger.log('エラー: フォルダIDがスクリプトプロパティに設定されていません。');
      SpreadsheetApp.getUi().alert('エラー', 'スクリプトプロパティにFOLDER_IDを設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
  
    // yyyy年mm月形式のサブフォルダ名を作成
    const now = new Date();
    const yearMonthFolderName = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy年MM月");
    const parentFolder = DriveApp.getFolderById(folderId);
    let subFolder;
    // サブフォルダが存在するか確認
    const folders = parentFolder.getFoldersByName(yearMonthFolderName);
    if (folders.hasNext()) {
      subFolder = folders.next();
    } else {
      subFolder = parentFolder.createFolder(yearMonthFolderName);
    }
  
    sheets.forEach(sheet => {
      const sheetName = sheet.getName(); // シート名 = 店舗名として利用
      // 実行時の月を取得（2025年7月1日なので「7」となる）
      const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M"); 
      const docTitle = `Google口コミ【${currentMonth}月／${sheetName}】`; // ドキュメントタイトル
  
      // 新しいGoogleドキュメントを作成
      const newDoc = DocumentApp.create(docTitle);
      const docFile = DriveApp.getFileById(newDoc.getId());
      docFile.moveTo(subFolder); // まずサブフォルダに移動
      const body = newDoc.getBody();
  
      // ドキュメントのタイトル/見出しを追加
      body.appendParagraph(`【${sheetName}】口コミデータ`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('---'); // 区切り線
      body.appendParagraph(''); // 空行で区切り
  
      // スプレッドシートのデータを読み込み（ヘッダー行を除いて2行目から開始）
      // 各行に1つの口コミデータが横に並んでいる想定
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
  
      if (lastRow <= 1 || lastColumn === 0) { // ヘッダー行のみの場合はデータなしと判断
        Logger.log(`シート「${sheetName}」には口コミデータがありません。スキップします。`);
        return;
      }
  
      // ヘッダー行は読み飛ばし、2行目から最終行までを各口コミとして処理
      // 取得範囲: 2行目, 1列目 から 最終行, 最終列 まで
      const allData = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  
      allData.forEach((row, index) => {
        // 各口コミの間に区切り線を追加（最初の口コミ以外）
        if (index > 0) {
          body.appendParagraph(''); // 空行で区切り
          body.appendParagraph('---'); // 区切り線
          body.appendParagraph(''); // 空行で区切り
        }
  
        // 各口コミの項目を取得 (スプレッドシートの列順を想定)
        const rawDate = row[0]; // 日付 (Dateオブジェクトまたは文字列)
        const rawRating = row[1]; // 評価の数字
        const name = String(row[2] || '').trim(); // 名前
        const content = String(row[3] || '').trim(); // 口コミ内容
  
        // 日付のフォーマットを「M月d日」に変換
        let formattedDate = String(rawDate || '').trim(); // デフォルトはrawDateを文字列化したもの
        if (rawDate instanceof Date) {
          formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "M月d日");
        } else if (typeof rawDate === 'string' && rawDate.match(/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/)) { // YYYY/MM/DD や YYYY-MM-DD 形式の文字列日付
          try {
            const parsedDate = new Date(rawDate);
            if (!isNaN(parsedDate.getTime())) { // 有効な日付にパースできた場合
              formattedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "M月d日");
            }
          } catch (e) {
            Logger.log(`日付のパースに失敗しました: ${rawDate}, エラー: ${e.message}`);
          }
        }
        
        // 評価を星の数に変換
        let starRating = '';
        const ratingNumber = parseInt(rawRating); // 評価を数値に変換
        if (!isNaN(ratingNumber) && ratingNumber >= 0 && ratingNumber <= 5) {
          starRating = '★'.repeat(ratingNumber) + '☆'.repeat(5 - ratingNumber);
        } else {
          starRating = String(rawRating || '(評価なし)'); // 無効な値の場合はそのまま表示または「評価なし」
        }
  
        // 各項目を改行して出力（ラベルなし）
        body.appendParagraph(`${formattedDate}`); // 日付のみ
        body.appendParagraph(`${starRating}`);   // 評価のみ
        // 名前は、元のドキュメントになかったため、今回は出力しない
        // body.appendParagraph(`${name}`); // 必要ならこの行を有効にする
        body.appendParagraph(`${content}`);     // 口コミ内容のみ
      });
  
      Logger.log(`ドキュメント「${newDoc.getName()}」を生成し、フォルダ「${subFolder.getName()}」に保存しました。URL: ${newDoc.getUrl()}`);
    });
  }