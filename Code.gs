/**
 * Gmail下書きを取得してスプレッドシートに保存するGoogle Apps Script
 * 
 * このスクリプトは以下の機能を提供します：
 * - Gmailの下書きメッセージを取得
 * - 下書きの情報をGoogle Sheetsに保存
 */

/**
 * メイン実行関数
 * この関数を実行してGmail下書きをスプレッドシートに保存します
 */
function main() {
  try {
    // スプレッドシートIDを設定（実際のスプレッドシートIDに変更してください）
    const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
    
    // スプレッドシートIDが設定されているかチェック
    if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
      console.error('エラー: スプレッドシートIDが設定されていません。');
      console.log('SPREADSHEET_ID変数に実際のスプレッドシートIDを設定してください。');
      return;
    }
    
    console.log('Gmail下書きの取得を開始します...');
    
    // Gmail下書きを取得
    const drafts = getDrafts();
    console.log(`${drafts.length}件の下書きを取得しました`);
    
    if (drafts.length > 0) {
      // スプレッドシートに保存
      saveToSheet(SPREADSHEET_ID, drafts);
      console.log('下書きの情報をスプレッドシートに保存しました');
    } else {
      console.log('保存する下書きがありません');
    }
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    console.error('エラーの詳細:', error.toString());
  }
}

/**
 * Gmail下書きを取得する関数
 * @return {Array} 下書き情報の配列
 */
function getDrafts() {
  try {
    // Gmail下書きを取得（最大100件）
    const drafts = GmailApp.getDrafts();
    const draftData = [];
    
    drafts.forEach((draft, index) => {
      try {
        const message = draft.getMessage();
        
        // 下書きの詳細情報を取得
        const draftInfo = {
          id: draft.getId(),
          subject: message.getSubject(),
          body: message.getPlainBody(),
          to: message.getTo(),
          cc: message.getCc(),
          bcc: message.getBcc(),
          date: message.getDate(),
          attachments: message.getAttachments().length,
          bodyPreview: message.getPlainBody().substring(0, 100) + '...' // 最初の100文字
        };
        
        draftData.push(draftInfo);
        
      } catch (error) {
        console.error(`下書き${index + 1}の処理中にエラーが発生しました:`, error);
      }
    });
    
    return draftData;
    
  } catch (error) {
    console.error('Gmail下書きの取得中にエラーが発生しました:', error);
    return [];
  }
}

/**
 * 下書きデータをスプレッドシートに保存する関数
 * @param {string} spreadsheetId - スプレッドシートのID
 * @param {Array} drafts - 保存する下書きデータ
 */
function saveToSheet(spreadsheetId, drafts) {
  try {
    console.log(`スプレッドシートID: ${spreadsheetId} への保存を開始します...`);
    
    // スプレッドシートIDの有効性をチェック
    if (!spreadsheetId || spreadsheetId === 'YOUR_SPREADSHEET_ID_HERE') {
      throw new Error('有効なスプレッドシートIDが設定されていません');
    }
    
    // スプレッドシートを開く
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      console.log(`スプレッドシート "${spreadsheet.getName()}" を開きました`);
    } catch (error) {
      console.error('スプレッドシートを開けませんでした:', error);
      console.log('スプレッドシートIDが正しいか、アクセス権限があるか確認してください。');
      throw new Error(`スプレッドシートにアクセスできません: ${error.message}`);
    }
    
    // シート名を設定（存在しない場合は作成）
    const sheetName = 'Gmail下書き';
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      console.log(`新しいシート "${sheetName}" を作成しました`);
    }
    
    // ヘッダー行を設定
    const headers = [
      '下書きID',
      '件名',
      '宛先',
      'CC',
      'BCC',
      '作成日時',
      '添付ファイル数',
      '本文プレビュー',
      '取得日時'
    ];
    
    // 既存のデータをクリア（ヘッダーは保持）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      console.log(`既存データ ${lastRow - 1} 行をクリアしました`);
    }
    
    // ヘッダーを設定（初回のみ）
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // ヘッダー行の書式設定
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      console.log('ヘッダー行を設定しました');
    }
    
    // 下書きデータを準備
    const currentTime = new Date();
    const rows = drafts.map(draft => [
      draft.id || '',
      draft.subject || '',
      draft.to || '',
      draft.cc || '',
      draft.bcc || '',
      draft.date || '',
      draft.attachments || 0,
      draft.bodyPreview || '',
      currentTime
    ]);
    
    // データをシートに書き込み
    if (rows.length > 0) {
      const dataRange = sheet.getRange(2, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // 列幅を自動調整
      sheet.autoResizeColumns(1, headers.length);
      
      console.log(`${rows.length}件のデータをシートに保存しました`);
    }
    
  } catch (error) {
    console.error('スプレッドシートへの保存中にエラーが発生しました:', error);
    console.error('エラーの詳細:', error.toString());
    throw error;
  }
}

/**
 * テスト実行関数：スプレッドシートIDを設定して下書きを取得するテスト
 * 実際のスプレッドシートIDを設定してから実行してください
 */
function testWithSpreadsheetId() {
  // TODO: ここに実際のスプレッドシートIDを設定してください
  const SPREADSHEET_ID = 'YOUR_ACTUAL_SPREADSHEET_ID';
  
  if (SPREADSHEET_ID === 'YOUR_ACTUAL_SPREADSHEET_ID') {
    console.log('スプレッドシートIDを設定してから実行してください');
    return;
  }
  
  main();
}

/**
 * 下書き数を確認する関数（テスト用）
 */
function checkDraftCount() {
  const drafts = GmailApp.getDrafts();
  console.log(`現在の下書き数: ${drafts.length}件`);
  return drafts.length;
}

/**
 * 新しいスプレッドシートを作成してテストする関数
 * この関数はスプレッドシートIDを自動で取得するため、手動設定が不要です
 */
function createTestSpreadsheet() {
  try {
    // 新しいスプレッドシートを作成
    const spreadsheet = SpreadsheetApp.create('Gmail下書き管理テスト');
    const spreadsheetId = spreadsheet.getId();
    
    console.log(`新しいスプレッドシートを作成しました:`);
    console.log(`名前: ${spreadsheet.getName()}`);
    console.log(`ID: ${spreadsheetId}`);
    console.log(`URL: ${spreadsheet.getUrl()}`);
    
    // Gmail下書きを取得
    const drafts = getDrafts();
    console.log(`${drafts.length}件の下書きを取得しました`);
    
    if (drafts.length > 0) {
      // スプレッドシートに保存
      saveToSheet(spreadsheetId, drafts);
      console.log('下書きの情報をスプレッドシートに保存しました');
      console.log(`スプレッドシートURL: ${spreadsheet.getUrl()}`);
    } else {
      console.log('保存する下書きがありません');
    }
    
    return {
      spreadsheetId: spreadsheetId,
      url: spreadsheet.getUrl(),
      name: spreadsheet.getName()
    };
    
  } catch (error) {
    console.error('テストスプレッドシート作成中にエラーが発生しました:', error);
    console.error('エラーの詳細:', error.toString());
    throw error;
  }
}

/**
 * スプレッドシートIDの有効性をテストする関数
 * @param {string} spreadsheetId - テストするスプレッドシートID
 */
function testSpreadsheetAccess(spreadsheetId) {
  try {
    console.log(`スプレッドシートID: ${spreadsheetId} のアクセステストを開始...`);
    
    if (!spreadsheetId) {
      console.error('スプレッドシートIDが指定されていません');
      return false;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log(`✅ スプレッドシートにアクセス成功:`);
    console.log(`   名前: ${spreadsheet.getName()}`);
    console.log(`   URL: ${spreadsheet.getUrl()}`);
    console.log(`   シート数: ${spreadsheet.getSheets().length}`);
    
    return true;
    
  } catch (error) {
    console.error('❌ スプレッドシートアクセステスト失敗:', error);
    console.log('確認事項:');
    console.log('1. スプレッドシートIDが正しいか');
    console.log('2. スプレッドシートにアクセス権限があるか');
    console.log('3. スプレッドシートが削除されていないか');
    return false;
  }
}
