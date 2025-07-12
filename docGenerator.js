/**
 * docGenerator.js - Google Apps Scriptによる口コミデータのドキュメント生成ツール
 * 
 * 指定した年月のデータを抽出し、見やすいドキュメントを生成します
 */

/**
 * 出力先フォルダIDをダイアログで入力し、スクリプトプロパティに保存します
 * @returns {boolean} 設定が成功した場合はtrue
 */
function selectOutputFolder() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('出力先フォルダIDを入力してください', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  
  const folderId = response.getResponseText().trim();
  if (!folderId) {
    ui.alert('エラー', 'フォルダIDが入力されていません。', ui.ButtonSet.OK);
    return false;
  }
  
  try {
    // フォルダIDの存在確認
    DriveApp.getFolderById(folderId);
    PropertiesService.getScriptProperties().setProperty('FOLDER_ID', folderId);
    ui.alert('成功', 'フォルダIDを保存しました。', ui.ButtonSet.OK);
    return true;
  } catch (error) {
    ui.alert('エラー', '無効なフォルダIDです: ' + error.message, ui.ButtonSet.OK);
    Logger.log('フォルダID設定エラー: ' + error);
    return false;
  }
}

/**
 * 年・月をプロンプトで入力し、プロパティに保存します
 * @returns {boolean} 設定が成功した場合はtrue
 */
function selectOutputMonth() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  
  // 現在の値を取得
  const currentYear = scriptProps.getProperty('TARGET_YEAR') || new Date().getFullYear();
  const currentMonth = scriptProps.getProperty('TARGET_MONTH') || (new Date().getMonth() + 1);
  
  // 年の入力
  const yearRes = ui.prompt(
    '出力する年を入力してください', 
    `現在の設定: ${currentYear}（空欄で変更しない）`, 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (yearRes.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  
  const yearInput = yearRes.getResponseText().trim();
  const year = yearInput ? yearInput : currentYear;
  
  // 年の妥当性チェック
  if (isNaN(year) || year < 2000 || year > 2100) {
    ui.alert('エラー', '有効な年を入力してください（2000〜2100）', ui.ButtonSet.OK);
    return false;
  }
  
  // 月の入力
  const monthRes = ui.prompt(
    '出力する月を入力してください', 
    `現在の設定: ${currentMonth}（空欄で変更しない）`, 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (monthRes.getSelectedButton() !== ui.Button.OK) {
    return false;
  }
  
  const monthInput = monthRes.getResponseText().trim();
  const month = monthInput ? monthInput : currentMonth;
  
  // 月の妥当性チェック
  if (isNaN(month) || month < 1 || month > 12) {
    ui.alert('エラー', '有効な月を入力してください（1〜12）', ui.ButtonSet.OK);
    return false;
  }
  
  // 値を保存
  scriptProps.setProperty('TARGET_YEAR', year);
  scriptProps.setProperty('TARGET_MONTH', month);
  
  ui.alert('成功', `出力対象を ${year}年${month}月 に設定しました。`, ui.ButtonSet.OK);
  return true;
}

/**
 * 日付文字列をフォーマットする
 * @param {Date|string} dateValue - 日付オブジェクトまたは日付文字列
 * @param {string} format - 出力形式 ('yyyy年M月'や'M月d日'など)
 * @returns {string} フォーマット済み日付文字列
 */
function formatDateString(dateValue, format) {
  let dateObj;
  
  if (dateValue instanceof Date) {
    dateObj = dateValue;
  } else if (typeof dateValue === 'string') {
    // YYYY-MM-DD または YYYY/MM/DD 形式の文字列をパース
    const match = String(dateValue).match(/^([0-9]{4})[\/\-]([0-9]{1,2})[\/\-]([0-9]{1,2})$/);
    if (match) {
      // JavaScriptのDateは月が0始まりなので-1する
      dateObj = new Date(match[1], match[2] - 1, match[3]);
    } else {
      return String(dateValue); // パースできなかった場合はそのまま返す
    }
  } else {
    return String(dateValue || ''); // 不明な型の場合はそのまま返す
  }
  
  // 日付として無効な場合
  if (isNaN(dateObj.getTime())) {
    return String(dateValue || '');
  }
  
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), format);
}

/**
 * 星評価を生成する
 * @param {number|string} rating - 評価値（0〜5）
 * @returns {string} 星記号で表現した評価（★★★☆☆ など）
 */
function generateStarRating(rating) {
  const ratingNumber = Number(rating);
  
  if (isNaN(ratingNumber) || ratingNumber < 0 || ratingNumber > 5) {
    return String(rating || '(評価なし)');
  }
  
  const filledStars = Math.round(ratingNumber);
  return '★'.repeat(filledStars) + '☆'.repeat(5 - filledStars);
}

/**
 * 指定した年月の口コミデータを抽出してドキュメントを生成します
 * 保存されているスクリプトプロパティから設定を読み込みます
 */
function generateDoc() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const spreadsheetId = scriptProps.getProperty('SPREADSHEET_ID');
    const folderId = scriptProps.getProperty('FOLDER_ID');
    const targetYear = scriptProps.getProperty('TARGET_YEAR');
    const targetMonth = scriptProps.getProperty('TARGET_MONTH');

    // 必須プロパティのチェック
    if (!spreadsheetId) {
      ui.alert('エラー', 'スプレッドシートIDがスクリプトプロパティに設定されていません。', ui.ButtonSet.OK);
      return;
    }
    if (!folderId) {
      ui.alert('エラー', 'フォルダIDがスクリプトプロパティに設定されていません。', ui.ButtonSet.OK);
      return;
    }

    // 対象年月の決定
    let y, m;
    if (targetYear && targetMonth) {
      y = parseInt(targetYear, 10);
      m = parseInt(targetMonth, 10);
    } else {
      const now = new Date();
      y = now.getFullYear();
      m = now.getMonth() + 1;
    }
    
    const yearMonthStr = `${y}年${m}月`;
    SpreadsheetApp.getActiveSpreadsheet().toast(`「${yearMonthStr}」のデータを出力します。`, '出力対象', 5);
    
    // 出力先フォルダ確認
    let parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(folderId);
    } catch (error) {
      ui.alert('エラー', `指定されたフォルダが見つかりません: ${error.message}`, ui.ButtonSet.OK);
      return;
    }
    
    // スプレッドシート読み込み
    let ss;
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      ui.alert('エラー', `スプレッドシートが開けません: ${error.message}`, ui.ButtonSet.OK);
      return;
    }
    
    // 全店舗まとめ用データを格納する配列
    const mergedContents = [];
    const sheets = ss.getSheets();
    
    // データ処理
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const lastRow = sheet.getLastRow();
      const lastColumn = sheet.getLastColumn();
      
      if (lastRow <= 1 || lastColumn === 0) {
        Logger.log(`シート「${sheetName}」にはデータがありません。スキップします。`);
        return;
      }
      
      try {
        const allData = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

        // 指定年月の口コミのみ抽出
        const filteredData = allData.filter(row => {
          const rawDate = row[0];
          let date = null;
          
          if (rawDate instanceof Date) {
            date = rawDate;
          } else if (typeof rawDate === 'string' && rawDate.trim()) {
            // YYYY-MM-DD または YYYY/MM/DD 形式の文字列をパース
            const match = String(rawDate).match(/^([0-9]{4})[\/\-]([0-9]{1,2})[\/\-]([0-9]{1,2})$/);
            if (match) {
              date = new Date(match[1], match[2] - 1, match[3]);
            }
          }
          
          if (!date || isNaN(date.getTime())) {
            return false;
          }
          
          return date.getFullYear() === y && (date.getMonth() + 1) === m;
        });

        if (filteredData.length === 0) {
          Logger.log(`シート「${sheetName}」に対象月(${yearMonthStr})の口コミデータがありません。スキップします。`);
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
          
          // 日付フォーマット
          const formattedDate = formatDateString(rawDate, "M月d日");
          
          // 評価の星表示
          const starRating = generateStarRating(rawRating);
          
          paragraphs.push(`${formattedDate}`);
          paragraphs.push(`${starRating}`);
          
          // 投稿者名があれば表示
          if (name) {
            paragraphs.push(`投稿者: ${name}`);
          }
          
          paragraphs.push('');
          paragraphs.push(`${content}`);
        });
        
        mergedContents.push({ shopName: sheetName, paragraphs });
        
      } catch (error) {
        Logger.log(`シート「${sheetName}」の処理中にエラーが発生しました: ${error}`);
      }
    });

    // --- 全店舗まとめドキュメント作成 ---
    if (mergedContents.length > 0) {
      try {
        const mergedDocTitle = `【${yearMonthStr}】全店舗口コミレポート`;
        const mergedDoc = DocumentApp.create(mergedDocTitle);
        const mergedDocFile = DriveApp.getFileById(mergedDoc.getId());
        mergedDocFile.moveTo(parentFolder);
        
        const mergedBody = mergedDoc.getBody();
        mergedBody.clear();
        
        // ドキュメント全体のフォント設定
        mergedBody.setFontFamily('Noto Sans');
        
        // 見出し1としてタイトルを追加
        mergedBody.appendParagraph(mergedDocTitle)
          .setHeading(DocumentApp.ParagraphHeading.HEADING1)
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
        // 各店舗のデータを追加
        mergedContents.forEach((item, idx) => {
          // 最初の店舗以外はページ区切りを入れる
          if (idx > 0) {
            mergedBody.appendPageBreak();
          }
          
          // 店舗データ追加
          item.paragraphs.forEach((p, i) => {
            let para;
            if (i === 0) {
              // 店舗名は見出し2として表示
              para = mergedBody.appendParagraph(p).setHeading(DocumentApp.ParagraphHeading.HEADING2);
            } else {
              para = mergedBody.appendParagraph(p);
              para.setFontSize(12);
            }
            para.setFontFamily('Noto Sans');
          });
        });
        
        // 成功メッセージ
        const docUrl = mergedDoc.getUrl();
        ui.alert('成功', `全店舗まとめドキュメント「${mergedDocTitle}」を生成しました。\n\nURL: ${docUrl}`, ui.ButtonSet.OK);
        Logger.log(`ドキュメント「${mergedDocTitle}」を作成しました。URL: ${docUrl}`);
        
      } catch (error) {
        ui.alert('エラー', `ドキュメント作成中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
        Logger.log(`ドキュメント作成エラー: ${error}`);
      }
    } else {
      ui.alert('情報', `対象期間(${yearMonthStr})のデータが見つかりませんでした。`, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    ui.alert('エラー', `処理中にエラーが発生しました: ${error.message}`, ui.ButtonSet.OK);
    Logger.log(`generateDoc エラー: ${error}`);
  }
}