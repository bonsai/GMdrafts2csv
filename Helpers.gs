/**
 * Gmail下書き管理のヘルパー関数集
 * 
 * このファイルには、メインスクリプトで使用する
 * 追加のヘルパー関数が含まれています。
 */

/**
 * 特定の条件に基づいて下書きをフィルタリングする関数
 * @param {Array} drafts - 下書きの配列
 * @param {Object} filters - フィルター条件
 * @return {Array} フィルタリング済みの下書き配列
 */
function filterDrafts(drafts, filters = {}) {
  return drafts.filter(draft => {
    // 件名でフィルタリング
    if (filters.subjectContains && 
        !draft.subject.toLowerCase().includes(filters.subjectContains.toLowerCase())) {
      return false;
    }
    
    // 日付でフィルタリング
    if (filters.dateAfter && draft.date < filters.dateAfter) {
      return false;
    }
    
    if (filters.dateBefore && draft.date > filters.dateBefore) {
      return false;
    }
    
    // 宛先でフィルタリング
    if (filters.toContains && 
        !draft.to.toLowerCase().includes(filters.toContains.toLowerCase())) {
      return false;
    }
    
    return true;
  });
}

/**
 * 下書きの統計情報を取得する関数
 * @param {Array} drafts - 下書きの配列
 * @return {Object} 統計情報
 */
function getDraftStatistics(drafts) {
  const stats = {
    totalCount: drafts.length,
    withAttachments: 0,
    withoutSubject: 0,
    averageBodyLength: 0,
    oldestDraft: null,
    newestDraft: null
  };
  
  if (drafts.length === 0) {
    return stats;
  }
  
  let totalBodyLength = 0;
  let oldestDate = new Date();
  let newestDate = new Date(0);
  
  drafts.forEach(draft => {
    // 添付ファイル付きの下書きをカウント
    if (draft.attachments > 0) {
      stats.withAttachments++;
    }
    
    // 件名なしの下書きをカウント
    if (!draft.subject || draft.subject.trim() === '') {
      stats.withoutSubject++;
    }
    
    // 本文の長さを累計
    totalBodyLength += draft.body.length;
    
    // 最古・最新の下書きを特定
    if (draft.date < oldestDate) {
      oldestDate = draft.date;
      stats.oldestDraft = draft;
    }
    
    if (draft.date > newestDate) {
      newestDate = draft.date;
      stats.newestDraft = draft;
    }
  });
  
  // 平均本文長を計算
  stats.averageBodyLength = Math.round(totalBodyLength / drafts.length);
  
  return stats;
}

/**
 * 下書きデータをCSV形式で出力する関数
 * @param {Array} drafts - 下書きの配列
 * @return {string} CSV形式の文字列
 */
function exportDraftsToCSV(drafts) {
  const headers = [
    '下書きID',
    '件名',
    '宛先',
    'CC',
    'BCC',
    '作成日時',
    '添付ファイル数',
    '本文プレビュー'
  ];
  
  let csv = headers.join(',') + '\n';
  
  drafts.forEach(draft => {
    const row = [
      `"${draft.id}"`,
      `"${draft.subject.replace(/"/g, '""')}"`,
      `"${draft.to.replace(/"/g, '""')}"`,
      `"${draft.cc.replace(/"/g, '""')}"`,
      `"${draft.bcc.replace(/"/g, '""')}"`,
      `"${draft.date.toISOString()}"`,
      draft.attachments,
      `"${draft.bodyPreview.replace(/"/g, '""')}"`
    ];
    csv += row.join(',') + '\n';
  });
  
  return csv;
}

/**
 * スプレッドシートに統計情報シートを作成する関数
 * @param {string} spreadsheetId - スプレッドシートのID
 * @param {Object} stats - 統計情報
 */
function createStatsSheet(spreadsheetId, stats) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = '下書き統計';
    
    // 既存の統計シートを削除
    const existingSheet = spreadsheet.getSheetByName(sheetName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    // 新しい統計シートを作成
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // 統計データを準備
    const data = [
      ['項目', '値'],
      ['総下書き数', stats.totalCount],
      ['添付ファイル付き', stats.withAttachments],
      ['件名なし', stats.withoutSubject],
      ['平均本文長', stats.averageBodyLength + ' 文字'],
      ['最古の下書き', stats.oldestDraft ? stats.oldestDraft.date.toLocaleDateString('ja-JP') : 'なし'],
      ['最新の下書き', stats.newestDraft ? stats.newestDraft.date.toLocaleDateString('ja-JP') : 'なし'],
      ['最後の更新', new Date().toLocaleString('ja-JP')]
    ];
    
    // データをシートに書き込み
    const range = sheet.getRange(1, 1, data.length, 2);
    range.setValues(data);
    
    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, 2);
    headerRange.setBackground('#34a853');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // 列幅を自動調整
    sheet.autoResizeColumns(1, 2);
    
    console.log('統計情報シートを作成しました');
    
  } catch (error) {
    console.error('統計情報シートの作成中にエラーが発生しました:', error);
  }
}

/**
 * 統計情報付きでデータを保存する関数
 * @param {string} spreadsheetId - スプレッドシートのID
 */
function saveWithStatistics(spreadsheetId) {
  try {
    // 下書きを取得
    const drafts = getDrafts();
    console.log(`${drafts.length}件の下書きを取得しました`);
    
    if (drafts.length > 0) {
      // メインデータを保存
      saveToSheet(spreadsheetId, drafts);
      
      // 統計情報を計算
      const stats = getDraftStatistics(drafts);
      console.log('統計情報:', stats);
      
      // 統計情報シートを作成
      createStatsSheet(spreadsheetId, stats);
      
      console.log('データと統計情報を保存しました');
    } else {
      console.log('保存する下書きがありません');
    }
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
  }
}

/**
 * 古い下書きを削除する関数（注意：実際に下書きが削除されます）
 * @param {number} daysOld - 何日前より古い下書きを削除するか
 * @param {boolean} dryRun - true の場合は削除せずにログ出力のみ
 */
function deleteOldDrafts(daysOld = 30, dryRun = true) {
  try {
    const drafts = GmailApp.getDrafts();
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysOld);
    
    let deleteCount = 0;
    
    drafts.forEach((draft, index) => {
      try {
        const message = draft.getMessage();
        const draftDate = message.getDate();
        
        if (draftDate < cutoffDate) {
          if (dryRun) {
            console.log(`削除対象: ${message.getSubject()} (${draftDate.toLocaleDateString('ja-JP')})`);
          } else {
            draft.delete();
            console.log(`削除しました: ${message.getSubject()}`);
          }
          deleteCount++;
        }
        
      } catch (error) {
        console.error(`下書き${index + 1}の処理中にエラーが発生しました:`, error);
      }
    });
    
    if (dryRun) {
      console.log(`${deleteCount}件の古い下書きが削除対象です（${daysOld}日以前）`);
      console.log('実際に削除するには dryRun を false に設定してください');
    } else {
      console.log(`${deleteCount}件の古い下書きを削除しました`);
    }
    
    return deleteCount;
    
  } catch (error) {
    console.error('古い下書きの削除中にエラーが発生しました:', error);
    return 0;
  }
}
