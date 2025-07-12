/**
 * 出力先フォルダIDをダイアログで入力し、スクリプトプロパティに保存
 */
function selectOutputFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('出力先フォルダIDを入力してください', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const folderId = response.getResponseText().trim();
    if (folderId) {
      PropertiesService.getScriptProperties().setProperty('FOLDER_ID', folderId);
      ui.alert('フォルダIDを保存しました。');
    } else {
      ui.alert('フォルダIDが入力されていません。');
    }
  }
}

/**
 * 年・月をプロンプトで入力し、空欄でなければプロパティに保存
 */
function selectOutputMonth() {
  const ui = SpreadsheetApp.getUi();
  const yearRes = ui.prompt('出力する年を入力してください（例: 2024、空欄で変更しない）', ui.ButtonSet.OK_CANCEL);
  if (yearRes.getSelectedButton() !== ui.Button.OK) return;
  const year = yearRes.getResponseText().trim();
  if (year) {
    PropertiesService.getScriptProperties().setProperty('TARGET_YEAR', year);
  }

  const monthRes = ui.prompt('出力する月を入力してください（例: 6、空欄で変更しない）', ui.ButtonSet.OK_CANCEL);
  if (monthRes.getSelectedButton() !== ui.Button.OK) return;
  const month = monthRes.getResponseText().trim();
  if (month) {
    PropertiesService.getScriptProperties().setProperty('TARGET_MONTH', month);
  }
}

/**
 * 指定した年月の口コミデータのみを出力する（将来的にUIから指定可能な設計）
 * @param {string} targetYearMonth - 'yyyy年M月' 形式（例: '2024年6月'）。未指定時は直前の月。
 */
function generateDoc() {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    const targetYear = PropertiesService.getScriptProperties().getProperty('TARGET_YEAR');
    const targetMonth = PropertiesService.getScriptProperties().getProperty('TARGET_MONTH');

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

    // 対象年月の決定
    let yearMonthStr;
    let y, m;
    if (targetYear && targetMonth) {
      yearMonthStr = `${targetYear}年${parseInt(targetMonth, 10)}月`;
      y = parseInt(targetYear, 10);
      m = parseInt(targetMonth, 10);
    } else {
      // 未指定なら今月
      const now = new Date();
      yearMonthStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy年M月");
      y = now.getFullYear();
      m = now.getMonth() + 1;
    }
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
  
    // 全店舗まとめ用データを格納する配列
    const mergedContents = [];
  
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
  
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      // 個別ドキュメント作成は行わず、内容のみ収集
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      if (lastRow <= 1 || lastColumn === 0) {
        Logger.log(`シート「${sheetName}」には口コミデータがありません。スキップします。`);
        return;
      }
      const allData = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

      // 指定年月の口コミのみ抽出
      const filteredData = allData.filter(row => {
        const rawDate = row[0];
        let yy = null, mm = null;
        if (rawDate instanceof Date) {
          yy = rawDate.getFullYear();
          mm = rawDate.getMonth() + 1;
        } else if (typeof rawDate === 'string' && rawDate.match(/^[0-9]{4}[\/\-][0-9]{1,2}[\/\-][0-9]{1,2}$/)) {
          const parsed = new Date(rawDate);
          if (!isNaN(parsed.getTime())) {
            yy = parsed.getFullYear();
            mm = parsed.getMonth() + 1;
          }
        }
        return yy === y && mm === m;
      });

      if (filteredData.length === 0) {
        Logger.log(`シート「${sheetName}」に対象月の口コミデータがありません。スキップします。`);
        return;
      }

      // 1店舗分の内容を配列で構築
      const paragraphs = [];
      paragraphs.push(`【${sheetName}】口コミデータ`);
      paragraphs.push('---');
      filteredData.forEach((row, index) => {
        if (index > 0) {
          paragraphs.push('');
          paragraphs.push('---');
        }
        const rawDate = row[0];
        const rawRating = row[1];
        const name = String(row[2] || '').trim();
        const content = String(row[3] || '').trim();
        let formattedDate = String(rawDate || '').trim();
        if (rawDate instanceof Date) {
          formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "M月d日");
        } else if (typeof rawDate === 'string' && rawDate.match(/^[0-9]{4}[\/\-][0-9]{1,2}[\/\-][0-9]{1,2}$/)) {
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
        paragraphs.push(`${formattedDate}`);
        paragraphs.push(`${starRating}`);
        paragraphs.push('');
        paragraphs.push(`${content}`);
      });
      mergedContents.push({ paragraphs });
    });

    // --- 全店舗まとめドキュメント作成 ---
    if (mergedContents.length > 0) {
      const mergedDocTitle = `【${yearMonthStr}】全店舗口コミレポート`;
      const mergedDoc = DocumentApp.create(mergedDocTitle);
      const mergedDocFile = DriveApp.getFileById(mergedDoc.getId());
      mergedDocFile.moveTo(subFolder);
      const mergedBody = mergedDoc.getBody();
      mergedBody.clear();
      // 全体のフォントをNoto Sansに統一
      mergedBody.setFontFamily('Noto Sans');
      mergedContents.forEach((item, idx) => {
        item.paragraphs.forEach((p, i) => {
          let para;
          if (i === 0) {
            para = mergedBody.appendParagraph(p).setHeading(DocumentApp.ParagraphHeading.HEADING2);
          } else {
            para = mergedBody.appendParagraph(p);
            para.setFontSize(12);
          }
          // 通常段落・見出しともにフォントをNoto Sansに統一
          para.setFontFamily('Noto Sans');
        });
        if (idx < mergedContents.length - 1) {
          mergedBody.appendPageBreak();
        }
      });
      Logger.log(`全店舗まとめドキュメント「${mergedDoc.getName()}」を生成し、フォルダ「${subFolder.getName()}」に保存しました。URL: ${mergedDoc.getUrl()}`);
    }
  }