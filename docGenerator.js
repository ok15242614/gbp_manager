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
  
    // 合体用データを格納する配列
    const mergedContents = [];
  
    sheets.forEach(sheet => {
      const sheetName = sheet.getName(); // シート名 = 店舗名として利用
      const currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M"); 
      const docTitle = `Google口コミ【${currentMonth}月／${sheetName}】`;
  
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
  
      // 合体用データとして保存（装飾を維持したい場合はgetBody().copy()も可）
      mergedContents.push({
        title: docTitle,
        paragraphs: body.getParagraphs().map(p => p.copy())
      });
  
      Logger.log(`ドキュメント「${newDoc.getName()}」を生成し、フォルダ「${subFolder.getName()}」に保存しました。URL: ${newDoc.getUrl()}`);
    });
  
    // --- 合体ドキュメント作成 ---
    if (mergedContents.length > 0) {
      const mergedDoc = DocumentApp.create(`合体ドキュメント_${yearMonthFolderName}`);
      const mergedDocFile = DriveApp.getFileById(mergedDoc.getId());
      mergedDocFile.moveTo(subFolder);
      const mergedBody = mergedDoc.getBody();
      mergedBody.clear(); // デフォルトの空段落を削除
      mergedContents.forEach((item, idx) => {
        mergedBody.appendParagraph(item.title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        item.paragraphs.forEach(p => mergedBody.appendParagraph(p));
        if (idx < mergedContents.length - 1) {
          mergedBody.appendPageBreak();
        }
      });
      Logger.log(`合体ドキュメント「${mergedDoc.getName()}」を生成し、フォルダ「${subFolder.getName()}」に保存しました。URL: ${mergedDoc.getUrl()}`);
    }
  }