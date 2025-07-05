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
  
    // 年月文字列を取得（例：2024年6月）
    const now = new Date();
    const yearMonthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy年M月");
    const yearMonthFolderName = yearMonthStr;
    const parentFolder = DriveApp.getFolderById(folderId);
    let subFolder;
    // サブフォルダが存在するか確認
    const folders = parentFolder.getFoldersByName(yearMonthFolderName);
    if (folders.hasNext()) {
      subFolder = folders.next();
    } else {
      subFolder = parentFolder.createFolder(yearMonthFolderName);
    }
  
    // 合体用データを格納する配列
    const mergedContents = [];
  
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const docTitle = `【${yearMonthStr}】${sheetName} 口コミレポート`;
  
      // 新しいGoogleドキュメントを作成
      const newDoc = DocumentApp.create(docTitle);
      const docFile = DriveApp.getFileById(newDoc.getId());
      docFile.moveTo(subFolder); // まずサブフォルダに移動
      const body = newDoc.getBody();
  
      // ドキュメントのタイトル/見出しを追加
      body.appendParagraph(`【${sheetName}】口コミデータ`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
      body.appendParagraph('---');
      body.appendParagraph('');
  
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      if (lastRow <= 1 || lastColumn === 0) {
        Logger.log(`シート「${sheetName}」には口コミデータがありません。スキップします。`);
        return;
      }
      const allData = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  
      allData.forEach((row, index) => {
        if (index > 0) {
          body.appendParagraph('');
          body.appendParagraph('---');
          body.appendParagraph('');
        }
        const rawDate = row[0];
        const rawRating = row[1];
        const name = String(row[2] || '').trim();
        const content = String(row[3] || '').trim();
        let formattedDate = String(rawDate || '').trim();
        if (rawDate instanceof Date) {
          formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "M月d日");
        } else if (typeof rawDate === 'string' && rawDate.match(/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/)) {
          try {
            const parsedDate = new Date(rawDate);
            if (!isNaN(parsedDate.getTime())) {
              formattedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "M月d日");
            }
          } catch (e) {
            Logger.log(`日付のパースに失敗しました: ${rawDate}, エラー: ${e.message}`);
          }
        }
        let starRating = '';
        const ratingNumber = parseInt(rawRating);
        if (!isNaN(ratingNumber) && ratingNumber >= 0 && ratingNumber <= 5) {
          starRating = '★'.repeat(ratingNumber) + '☆'.repeat(5 - ratingNumber);
        } else {
          starRating = String(rawRating || '(評価なし)');
        }
        body.appendParagraph(`${formattedDate}`);
        body.appendParagraph(`${starRating}`);
        body.appendParagraph(`${content}`);
      });
  
      // 合体用データとして保存（タイトルなし、本文のみ）
      mergedContents.push({
        paragraphs: body.getParagraphs().map(p => p.copy())
      });
  
      Logger.log(`ドキュメント「${newDoc.getName()}」を生成し、フォルダ「${subFolder.getName()}」に保存しました。URL: ${newDoc.getUrl()}`);
    });
  
    // --- 合体ドキュメント作成 ---
    if (mergedContents.length > 0) {
      const mergedDocTitle = `【${yearMonthStr}】全店舗口コミレポート`;
      const mergedDoc = DocumentApp.create(mergedDocTitle);
      const mergedDocFile = DriveApp.getFileById(mergedDoc.getId());
      mergedDocFile.moveTo(subFolder);
      const mergedBody = mergedDoc.getBody();
      mergedBody.clear(); // デフォルトの空段落を削除
      mergedContents.forEach((item, idx) => {
        item.paragraphs.forEach(p => mergedBody.appendParagraph(p));
        if (idx < mergedContents.length - 1) {
          mergedBody.appendPageBreak();
        }
      });
      Logger.log(`全店舗まとめドキュメント「${mergedDoc.getName()}」を生成し、フォルダ「${subFolder.getName()}」に保存しました。URL: ${mergedDoc.getUrl()}`);
    }
  }