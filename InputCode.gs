/**
 * 訪問看護 自動スケジューリング - 入力Webアプリ（Project B）
 * 管理者専用：マスタデータの表示（将来的には編集機能を追加）
 */

// スクリプトプロパティから設定を取得
const SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');
const ADMIN_SHEET = '管理者';
const OUTPUT_APP_URL = PropertiesService.getScriptProperties().getProperty('OUTPUT_APP_URL') || '';

// ============================================================
// Webアプリ エントリポイント
// ============================================================

/**
 * Webアプリのエントリポイント
 * 管理者判定を行い、適切なHTMLを返す
 */
function doGet() {
  try {
    // 設定チェック
    if (!SS_ID) {
      return HtmlService.createHtmlOutput(
        '<h2>設定エラー</h2><p>SS_ID が設定されていません。スクリプトプロパティを確認してください。</p>'
      ).setTitle('設定エラー');
    }

    // ユーザーメール取得
    const email = Session.getActiveUser().getEmail();

    // メールが取得できない場合
    if (!email) {
      return HtmlService.createHtmlOutputFromFile('NoAccess')
        .setTitle('アクセス権限エラー')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // 管理者判定
    if (!isAdmin_(email)) {
      return HtmlService.createHtmlOutputFromFile('NoAccess')
        .setTitle('アクセス権限エラー')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // 管理者：入力画面を表示
    return HtmlService.createHtmlOutputFromFile('InputIndex')
      .setTitle('訪問看護 自動スケジューリング - 入力管理')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (e) {
    console.error('doGet error:', e);
    return HtmlService.createHtmlOutput(
      '<h2>エラー</h2><p>' + e.message + '</p>'
    ).setTitle('エラー');
  }
}

// ============================================================
// 管理者判定
// ============================================================

/**
 * 管理者かどうかを判定
 * @param {string} email - ユーザーのメールアドレス
 * @returns {boolean} 管理者ならtrue
 */
function isAdmin_(email) {
  if (!email) return false;

  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(ADMIN_SHEET);

    if (!sheet) {
      console.warn('管理者シートが見つかりません: ' + ADMIN_SHEET);
      return false;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return false; // ヘッダーのみ

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf('email');
    const enabledIdx = headers.indexOf('enabled');

    if (emailIdx < 0) {
      console.warn('管理者シートにemail列がありません');
      return false;
    }

    const emailLower = email.toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEmail = String(row[emailIdx] || '').trim().toLowerCase();

      if (rowEmail === emailLower) {
        // enabled列がある場合はチェック
        if (enabledIdx >= 0) {
          const enabled = row[enabledIdx];
          // FALSE の場合のみ無効、空やTRUEは有効
          if (enabled === false || String(enabled).toUpperCase() === 'FALSE') {
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

/**
 * クライアントから呼び出し可能な管理者チェック
 */
function checkAdminStatus() {
  const email = Session.getActiveUser().getEmail();
  return {
    isAdmin: isAdmin_(email),
    email: email || '(取得不可)'
  };
}

// ============================================================
// シートデータ取得API
// ============================================================

/**
 * 汎用シートデータ取得（入力Webアプリ用）
 * @param {string} sheetName - シート名
 * @param {Object} opt - オプション
 * @param {number} opt.limit - 行数上限（デフォルト1000）
 */
function getSheetTable(sheetName, opt) {
  opt = opt || {};

  // 管理者チェック（APIレベルでも保護）
  const email = Session.getActiveUser().getEmail();
  if (!isAdmin_(email)) {
    throw new Error('権限がありません');
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return { headers: [], rows: [], rowCount: 0 };
  }

  const headers = values[0];
  let rows = values.slice(1);

  // 空行を除去（すべてのセルが空の行）
  rows = rows.filter(r => r.some(cell => cell !== '' && cell !== null && cell !== undefined));

  // 行数制限
  const limit = opt.limit || 1000;
  if (rows.length > limit) {
    rows = rows.slice(rows.length - limit);
  }

  // Date型をISO文字列に変換（JSON転送用）
  rows = rows.map(row => row.map(cell => {
    if (cell instanceof Date) {
      return { _type: 'date', value: cell.toISOString() };
    }
    return cell;
  }));

  return {
    headers,
    rows,
    rowCount: rows.length,
    sheetName
  };
}

/**
 * 利用可能なシート一覧を取得
 */
function getAvailableSheets() {
  const email = Session.getActiveUser().getEmail();
  if (!isAdmin_(email)) {
    throw new Error('権限がありません');
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheets = ss.getSheets();

  return sheets.map(s => ({
    name: s.getName(),
    rowCount: s.getLastRow(),
    colCount: s.getLastColumn()
  }));
}

// ============================================================
// 設定チェック
// ============================================================

/**
 * 設定状態をチェック
 */
function checkConfiguration() {
  const issues = [];

  if (!SS_ID) {
    issues.push('SS_ID が設定されていません');
  }

  try {
    const ss = SpreadsheetApp.openById(SS_ID);

    // 管理者シートチェック
    const adminSheet = ss.getSheetByName(ADMIN_SHEET);
    if (!adminSheet) {
      issues.push('「' + ADMIN_SHEET + '」シートが見つかりません');
    }

    // 入力シートチェック
    const inputSheets = ['患者マスタ', 'スタッフマスタ', '定期リクエスト', '個別変更リクエスト'];
    inputSheets.forEach(name => {
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
    issues
  };
}

/**
 * アプリリンク情報を取得
 */
function getAppLinks() {
  return {
    outputAppUrl: OUTPUT_APP_URL || null,
    hasOutputApp: !!OUTPUT_APP_URL
  };
}
