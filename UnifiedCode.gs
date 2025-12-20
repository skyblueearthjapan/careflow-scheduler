/**
 * 訪問看護 自動スケジューリング - 統合Webアプリ
 * URL1つで出力画面・入力画面を切り替え
 * 入力画面は管理者のみアクセス可能
 */

// ============================================================
// 定数（シート名・スプレッドシートID）
// ============================================================
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');

// シート名（typo防止用に定数化）
const SHEETS = {
  // 出力系
  WEEK_VIEW: '週ビュー',
  WEEKLY_REQUEST: '週間リクエスト',
  ASSIGN_RESULT: '割当結果',
  ASSIGN_NG: '割当不可',
  ROUTE_SUMMARY: 'ルートサマリ',
  // 入力系
  PATIENT_MASTER: '患者マスタ',
  STAFF_MASTER: 'スタッフマスタ',
  CHANGE_REQUEST: '個別変更リクエスト',
  // 権限
  ADMIN: '管理者',
  // その他
  LOG: '実行ログ'
};

// 出力タブ一覧
const OUTPUT_TABS = [
  { key: 'weekView', name: '週ビュー', sheetName: SHEETS.WEEK_VIEW },
  { key: 'weeklyRequest', name: '週間リクエスト', sheetName: SHEETS.WEEKLY_REQUEST },
  { key: 'assignResult', name: '割当結果', sheetName: SHEETS.ASSIGN_RESULT },
  { key: 'assignNg', name: '割当不可', sheetName: SHEETS.ASSIGN_NG },
  { key: 'routeSummary', name: 'ルートサマリ', sheetName: SHEETS.ROUTE_SUMMARY }
];

// 入力タブ一覧
const INPUT_TABS = [
  { key: 'patient', name: '患者マスタ', sheetName: SHEETS.PATIENT_MASTER },
  { key: 'staff', name: 'スタッフマスタ', sheetName: SHEETS.STAFF_MASTER },
  { key: 'change', name: '個別変更リクエスト', sheetName: SHEETS.CHANGE_REQUEST }
];

// 実行ボタン→関数名のマッピング（ホワイトリスト）
const JOB_MAP = {
  'weeklyRequest': '週間リクエストを生成_',
  'assignResult': '割当結果を作成_',
  'updateWeekView': '週ビューを更新_',
  'routeSummary': 'ルートサマリを作成_',
  'updateGeo': '位置情報を更新_'
};

// ============================================================
// Webアプリ エントリポイント
// ============================================================

/**
 * Webアプリのエントリポイント
 * @param {Object} e - イベントパラメータ
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'output';

  try {
    // 設定チェック
    if (!SPREADSHEET_ID) {
      return createErrorPage('設定エラー', 'SS_ID が設定されていません。スクリプトプロパティを確認してください。');
    }

    if (page === 'input') {
      // 入力ページは管理者のみ
      const email = Session.getActiveUser().getEmail();
      if (!email || !isAdmin_(email)) {
        return HtmlService.createHtmlOutputFromFile('UnifiedNoAccess')
          .setTitle('アクセス権限エラー')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
      return HtmlService.createHtmlOutputFromFile('UnifiedInput')
        .setTitle('訪問看護 自動スケジューリング - 入力管理')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // 出力ページ（デフォルト）
    return HtmlService.createHtmlOutputFromFile('UnifiedOutput')
      .setTitle('訪問看護 自動スケジューリング')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    console.error('doGet error:', err);
    return createErrorPage('エラー', err.message);
  }
}

/**
 * エラーページを作成
 */
function createErrorPage(title, message) {
  return HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;padding:40px;text-align:center;">' +
    '<h2 style="color:#E88B8B;">' + title + '</h2>' +
    '<p>' + message + '</p></body></html>'
  ).setTitle(title);
}

// ============================================================
// ユーティリティAPI（クライアントから呼び出し可能）
// ============================================================

/**
 * ベースURLを取得
 */
function getBaseUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * 現在のユーザーメールを取得
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail() || '';
}

/**
 * 現在のユーザーが管理者かどうか
 */
function checkIsAdmin() {
  const email = Session.getActiveUser().getEmail();
  return {
    isAdmin: email ? isAdmin_(email) : false,
    email: email || '(取得不可)'
  };
}

/**
 * 出力タブ一覧を取得
 */
function listOutputTabs() {
  return OUTPUT_TABS;
}

/**
 * 入力タブ一覧を取得
 */
function listInputTabs() {
  return INPUT_TABS;
}

// ============================================================
// 管理者判定
// ============================================================

/**
 * 管理者かどうかを判定
 * @param {string} email - ユーザーのメールアドレス
 * @returns {boolean}
 */
function isAdmin_(email) {
  if (!email) return false;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.ADMIN);
    if (!sheet) return false;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return false;

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const enabledIdx = headers.indexOf('enabled');
    const roleIdx = headers.indexOf('role');

    if (emailIdx < 0) return false;

    const emailLower = email.toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[emailIdx] || '').trim().toLowerCase();

      if (rowEmail === emailLower) {
        // enabled チェック
        if (enabledIdx >= 0) {
          const enabled = row[enabledIdx];
          if (enabled === false || String(enabled).toUpperCase() === 'FALSE') {
            return false;
          }
        }
        // role チェック（列がある場合のみ）
        if (roleIdx >= 0) {
          const role = String(row[roleIdx] || '').trim().toLowerCase();
          if (role !== 'admin') {
            return false;
          }
        }
        return true;
      }
    }
    return false;
  } catch (e) {
    console.error('isAdmin_ error:', e);
    return false;
  }
}

// ============================================================
// データ取得API（表示用）
// ============================================================

/**
 * 週ビューデータを取得
 */
function getWeekViewData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.WEEK_VIEW);
    if (!sheet) {
      throw new Error('シート「' + SHEETS.WEEK_VIEW + '」が見つかりません');
    }

    const lastRow = findLastDataRow_(sheet);
    const lastCol = 8; // A〜H列（職員名+7日分）

    if (lastRow < 1) {
      return { headerRow: [], bodyRows: [], rowCount: 0 };
    }

    const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headerRow = values[0] || [];
    const bodyRows = values.slice(1).filter(row => row[0] && String(row[0]).trim() !== '');

    return {
      headerRow: headerRow,
      bodyRows: bodyRows,
      rowCount: bodyRows.length
    };
  } catch (e) {
    console.error('getWeekViewData error:', e);
    throw e;
  }
}

/**
 * 汎用シートテーブルデータを取得
 * @param {string} sheetName - シート名
 * @param {number} limitRows - 行数上限（null/undefined なら全件）
 */
function getSheetTableData(sheetName, limitRows) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { header: [], rows: [], rowCount: 0, error: 'シートが見つかりません: ' + sheetName };
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length === 0) {
      return { header: [], rows: [], rowCount: 0 };
    }

    const header = data[0];
    let rows = data.slice(1).filter(row => row.some(cell => cell !== ''));

    // 行数制限
    const limit = limitRows || 500;
    if (rows.length > limit) {
      rows = rows.slice(rows.length - limit);
    }

    return {
      header: header,
      rows: rows,
      rowCount: rows.length,
      sheetName: sheetName
    };
  } catch (e) {
    console.error('getSheetTableData error:', e);
    return { header: [], rows: [], rowCount: 0, error: e.message };
  }
}

/**
 * 最終データ行を取得（A列基準）
 */
function findLastDataRow_(sheet) {
  const max = sheet.getMaxRows();
  if (max <= 1) return 1;
  const colA = sheet.getRange(1, 1, max, 1).getValues();
  let last = 1;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || '').trim() !== '') {
      last = i + 1;
    }
  }
  return last;
}

// ============================================================
// 実行API（GAS実行ボタン用）
// ============================================================

/**
 * ジョブを実行（ホワイトリスト制御）
 * @param {string} jobKey - ジョブキー
 */
function runJob(jobKey) {
  const startTime = new Date();

  // ホワイトリストチェック
  const funcName = JOB_MAP[jobKey];
  if (!funcName) {
    return { ok: false, message: '不明なジョブキー: ' + jobKey };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 関数を呼び出し
    let result;
    switch (jobKey) {
      case 'weeklyRequest':
        result = 週間リクエストを生成_(ss);
        break;
      case 'assignResult':
        result = 割当結果を作成_(ss);
        break;
      case 'updateWeekView':
        result = 週ビューを更新_(ss);
        break;
      case 'routeSummary':
        result = ルートサマリを作成_(ss);
        break;
      case 'updateGeo':
        result = 位置情報を更新_(ss);
        break;
      default:
        throw new Error('未実装のジョブ: ' + jobKey);
    }

    const message = (result && result.message) || '処理が完了しました';
    logExecution_(jobKey, true, message, startTime);
    return { ok: true, message: message };

  } catch (e) {
    console.error('runJob error:', e);
    logExecution_(jobKey, false, e.message, startTime);
    return { ok: false, message: e.message };
  }
}

/**
 * 実行ログを記録
 */
function logExecution_(action, success, message, startTime) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEETS.LOG);
    if (!sheet) {
      sheet = ss.insertSheet(SHEETS.LOG);
      sheet.appendRow(['タイムスタンプ', 'アクション', '成功', 'メッセージ', '実行時間(秒)', 'ユーザー']);
    }

    const endTime = new Date();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    const user = Session.getActiveUser().getEmail() || 'unknown';

    sheet.appendRow([
      endTime.toISOString(),
      action,
      success ? 'OK' : 'NG',
      message,
      duration,
      user
    ]);
  } catch (e) {
    console.error('logExecution_ error:', e);
  }
}

/**
 * 実行ログを取得
 * @param {number} limit - 取得件数
 */
function getExecutionLogs(limit) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.LOG);
    if (!sheet) {
      return { success: true, logs: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, logs: [] };
    }

    const rows = data.slice(1).reverse().slice(0, limit || 10);
    const logs = rows.map(row => ({
      timestamp: row[0],
      action: row[1],
      success: row[2] === 'OK',
      message: row[3],
      duration: row[4],
      user: row[5]
    }));

    return { success: true, logs: logs };
  } catch (e) {
    console.error('getExecutionLogs error:', e);
    return { success: false, error: e.message, logs: [] };
  }
}

// ============================================================
// 設定チェック
// ============================================================

/**
 * 設定状態をチェック
 */
function checkConfiguration() {
  const issues = [];

  if (!SPREADSHEET_ID) {
    issues.push('SS_ID が設定されていません');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 主要シートチェック
    const requiredSheets = [SHEETS.WEEK_VIEW, SHEETS.PATIENT_MASTER, SHEETS.STAFF_MASTER];
    requiredSheets.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) {
        issues.push('「' + name + '」シートが見つかりません');
      }
    });

  } catch (e) {
    issues.push('スプレッドシートにアクセスできません: ' + e.message);
  }

  return {
    success: issues.length === 0,
    issues: issues
  };
}
