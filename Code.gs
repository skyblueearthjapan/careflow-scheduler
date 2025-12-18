/**
 * 訪問看護 自動スケジューリング - Webアプリ
 * 第1段階：週ビュー表示 + GAS実行ボタン群
 */

// スクリプトプロパティから設定を取得
const SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');
const WEEKVIEW_SHEET = PropertiesService.getScriptProperties().getProperty('SHEET_WEEKVIEW') || '週ビュー';
const LOG_SHEET = PropertiesService.getScriptProperties().getProperty('SHEET_LOG') || '実行ログ';

/**
 * Webアプリのエントリポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('訪問看護 自動スケジューリング')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// 週ビュー取得API
// ============================================================

/**
 * 週ビューを取得（A列の最終スタッフ行まで）
 * @returns {Object} ヘッダ、行データ、メタ情報
 */
function getWeekView() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(WEEKVIEW_SHEET);

    if (!sheet) {
      throw new Error(`シート「${WEEKVIEW_SHEET}」が見つかりません`);
    }

    const lastRow = findLastStaffRow_(sheet);
    const lastCol = 8; // A〜H固定

    const values = sheet.getRange(1, 1, Math.max(1, lastRow), lastCol).getValues();

    // 1行目がヘッダ: [職員名, 12/15(Mon), ...]
    const header = values[0];

    // 2行目以降がスタッフ
    const rows = values.slice(1).map(r => ({
      staff: r[0],
      cells: r.slice(1) // B〜H
    })).filter(x => String(x.staff || '').trim() !== '');

    return {
      success: true,
      header,
      rows,
      meta: { lastRow, timestamp: new Date().toISOString() }
    };
  } catch (e) {
    console.error('getWeekView error:', e);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * A列を見て「2行目以降で最後に値がある行」を返す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {number} 最終行番号
 */
function findLastStaffRow_(sheet) {
  const max = sheet.getMaxRows();
  if (max <= 1) return 1;

  const colA = sheet.getRange(2, 1, max - 1, 1).getValues();
  let last = 1; // ヘッダ行だけの場合は1

  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || '').trim() !== '') {
      last = i + 2;
    }
  }
  return last;
}

// ============================================================
// GAS実行ラッパー関数（UIから呼ばれる）
// ============================================================

/**
 * 週間リクエストを生成
 */
function runGenerateWeeklyRequests() {
  const startTime = new Date();
  try {
    // TODO: 既存の週間リクエスト生成ロジックを呼び出す
    // generateWeeklyRequests_();

    const message = '週間リクエストの生成が完了しました';
    logExecution_('週間リクエスト生成', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    const message = e.message;
    logExecution_('週間リクエスト生成', false, message, startTime);
    return { success: false, error: message };
  }
}

/**
 * 割当結果を作成
 */
function runCreateAssignments() {
  const startTime = new Date();
  try {
    // TODO: 既存の割当結果作成ロジックを呼び出す
    // createAssignments_();

    const message = '割当結果の作成が完了しました';
    logExecution_('割当結果作成', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    const message = e.message;
    logExecution_('割当結果作成', false, message, startTime);
    return { success: false, error: message };
  }
}

/**
 * 週ビューを更新
 */
function runUpdateWeekView() {
  const startTime = new Date();
  try {
    // TODO: 既存の週ビュー更新ロジックを呼び出す
    // updateWeekView_();

    const message = '週ビューの更新が完了しました';
    logExecution_('週ビュー更新', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    const message = e.message;
    logExecution_('週ビュー更新', false, message, startTime);
    return { success: false, error: message };
  }
}

/**
 * ルートサマリを作成
 */
function runCreateRouteSummary() {
  const startTime = new Date();
  try {
    // TODO: 既存のルートサマリ作成ロジックを呼び出す
    // createRouteSummary_();

    const message = 'ルートサマリの作成が完了しました';
    logExecution_('ルートサマリ作成', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    const message = e.message;
    logExecution_('ルートサマリ作成', false, message, startTime);
    return { success: false, error: message };
  }
}

/**
 * 位置情報を更新
 */
function runUpdateLocation() {
  const startTime = new Date();
  try {
    // TODO: 既存の位置情報更新ロジックを呼び出す
    // updateLocation_();

    const message = '位置情報の更新が完了しました';
    logExecution_('位置情報更新', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    const message = e.message;
    logExecution_('位置情報更新', false, message, startTime);
    return { success: false, error: message };
  }
}

// ============================================================
// ログ機能
// ============================================================

/**
 * 実行ログをシートに記録
 * @param {string} actionName 実行ボタン名
 * @param {boolean} success 成否
 * @param {string} message メッセージ
 * @param {Date} startTime 開始時刻
 */
function logExecution_(actionName, success, message, startTime) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let logSheet = ss.getSheetByName(LOG_SHEET);

    // ログシートがなければ作成
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET);
      logSheet.appendRow(['実行日時', '実行者', '処理名', '成否', 'メッセージ', '実行時間(ms)']);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    }

    const endTime = new Date();
    const duration = endTime.getTime() - startTime.getTime();
    const user = Session.getActiveUser().getEmail() || '不明';

    logSheet.appendRow([
      Utilities.formatDate(endTime, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
      user,
      actionName,
      success ? '成功' : '失敗',
      message,
      duration
    ]);
  } catch (e) {
    console.error('logExecution_ error:', e);
  }
}

/**
 * 実行ログを取得（最新N件）
 * @param {number} limit 取得件数
 * @returns {Array} ログ配列
 */
function getExecutionLogs(limit = 20) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const logSheet = ss.getSheetByName(LOG_SHEET);

    if (!logSheet) {
      return { success: true, logs: [] };
    }

    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, logs: [] };
    }

    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    const data = logSheet.getRange(startRow, 1, numRows, 6).getValues();

    const logs = data.reverse().map(row => ({
      timestamp: row[0],
      user: row[1],
      action: row[2],
      success: row[3] === '成功',
      message: row[4],
      duration: row[5]
    }));

    return { success: true, logs };
  } catch (e) {
    console.error('getExecutionLogs error:', e);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ユーティリティ
// ============================================================

/**
 * スプレッドシートIDが設定されているか確認
 */
function checkConfiguration() {
  const issues = [];

  if (!SS_ID) {
    issues.push('SS_ID（スプレッドシートID）が設定されていません');
  }

  if (SS_ID) {
    try {
      const ss = SpreadsheetApp.openById(SS_ID);
      const sheet = ss.getSheetByName(WEEKVIEW_SHEET);
      if (!sheet) {
        issues.push(`シート「${WEEKVIEW_SHEET}」が見つかりません`);
      }
    } catch (e) {
      issues.push(`スプレッドシートにアクセスできません: ${e.message}`);
    }
  }

  return {
    success: issues.length === 0,
    issues,
    config: {
      SS_ID: SS_ID ? '設定済み' : '未設定',
      WEEKVIEW_SHEET,
      LOG_SHEET
    }
  };
}
