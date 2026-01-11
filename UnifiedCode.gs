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
  STAFF_CHANGE_REQUEST: 'スタッフ個別変更リクエスト',
  EVENT_REQUEST: 'イベントリクエスト',
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
  { key: 'change', name: '個別変更リクエスト', sheetName: SHEETS.CHANGE_REQUEST },
  { key: 'staffChange', name: 'スタッフ個別変更', sheetName: SHEETS.STAFF_CHANGE_REQUEST },
  { key: 'event', name: 'イベントリクエスト', sheetName: SHEETS.EVENT_REQUEST }
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
// 入力画面用API（CRUD操作）
// ============================================================

/**
 * 管理者チェック（APIで使用）
 * @returns {boolean}
 */
function requireAdmin_() {
  const email = Session.getActiveUser().getEmail();
  if (!email || !isAdmin_(email)) {
    throw new Error('権限がありません。管理者としてログインしてください。');
  }
  return true;
}

/**
 * 入力対象シート一覧を取得
 */
function input_listTables() {
  return INPUT_TABS;
}

/**
 * 入力用テーブルデータを取得（管理者チェック付き）
 * @param {string} sheetName - シート名
 */
function input_getTable(sheetName) {
  requireAdmin_();
  return getSheetTableData(sheetName, 1000);
}

/**
 * 行を末尾に追加
 * @param {string} sheetName - シート名
 * @param {Array<Array>} rows - 追加する行データ [[col1, col2, ...], ...]
 */
function input_appendRows(sheetName, rows) {
  requireAdmin_();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const numCols = headerRow.length;

    // 行データの列数を調整
    const normalizedRows = rows.map(row => {
      const newRow = new Array(numCols).fill('');
      for (let i = 0; i < Math.min(row.length, numCols); i++) {
        newRow[i] = row[i];
      }
      return newRow;
    });

    if (normalizedRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, normalizedRows.length, numCols).setValues(normalizedRows);
    }

    return { success: true, message: normalizedRows.length + ' 行を追加しました', addedCount: normalizedRows.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 行を削除（複数対応、行番号は1-based）
 * @param {string} sheetName - シート名
 * @param {Array<number>} rowNumbers - 削除する行番号（シートの行番号、1-based）
 */
function input_deleteRows(sheetName, rowNumbers) {
  requireAdmin_();

  if (!rowNumbers || rowNumbers.length === 0) {
    throw new Error('削除する行を指定してください');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    // 降順にソート（行ずれ防止）
    const sortedRows = rowNumbers.slice().sort((a, b) => b - a);

    // ヘッダー行（1行目）は削除不可
    const validRows = sortedRows.filter(r => r > 1);

    let deletedCount = 0;
    validRows.forEach(rowNum => {
      if (rowNum <= sheet.getLastRow()) {
        sheet.deleteRow(rowNum);
        deletedCount++;
      }
    });

    return { success: true, message: deletedCount + ' 行を削除しました', deletedCount: deletedCount };
  } finally {
    lock.releaseLock();
  }
}

/**
 * セルを更新
 * @param {string} sheetName - シート名
 * @param {Array<Object>} updates - 更新データ [{row: number, col: number, value: any}, ...]
 *                                   row/col は 1-based（シートの実座標）
 */
function input_updateCells(sheetName, updates) {
  requireAdmin_();

  if (!updates || updates.length === 0) {
    return { success: true, message: '更新するデータがありません', updatedCount: 0 };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    // ヘッダー行（1行目）は更新不可
    const validUpdates = updates.filter(u => u.row > 1);

    validUpdates.forEach(u => {
      sheet.getRange(u.row, u.col).setValue(u.value);
    });

    return { success: true, message: validUpdates.length + ' セルを更新しました', updatedCount: validUpdates.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 行をコピー（選択行を複製して末尾に追加）
 * @param {string} sheetName - シート名
 * @param {Array<number>} rowNumbers - コピー元の行番号（シートの行番号、1-based）
 */
function input_copyRows(sheetName, rowNumbers) {
  requireAdmin_();

  if (!rowNumbers || rowNumbers.length === 0) {
    throw new Error('コピーする行を指定してください');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    const lastCol = sheet.getLastColumn();
    const rowsToCopy = [];

    // 昇順にソート
    const sortedRows = rowNumbers.slice().sort((a, b) => a - b);

    sortedRows.forEach(rowNum => {
      if (rowNum > 1 && rowNum <= sheet.getLastRow()) {
        const rowData = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
        rowsToCopy.push(rowData);
      }
    });

    if (rowsToCopy.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToCopy.length, lastCol).setValues(rowsToCopy);
    }

    return { success: true, message: rowsToCopy.length + ' 行をコピーしました', copiedCount: rowsToCopy.length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 指定行の下に空行を挿入（ID自動採番付き）
 * @param {string} sheetName - シート名
 * @param {number} baseRowIndex - 基準行のシート行番号（1-based）。この行の下に挿入
 * @returns {Object} { success, newRowIndex, newRowData }
 */
function input_insertRowBelow(sheetName, baseRowIndex) {
  requireAdmin_();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    const numCols = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    const header = sheet.getRange(1, 1, 1, numCols).getValues()[0];

    // 挿入位置を決定
    let insertAt;
    if (!baseRowIndex || baseRowIndex < 2) {
      // 行選択なし or ヘッダー行 → 末尾に追加
      insertAt = lastRow + 1;
    } else if (baseRowIndex >= lastRow) {
      // 最終行選択 → その下に追加
      insertAt = lastRow + 1;
    } else {
      // 中間行選択 → その下に挿入
      insertAt = baseRowIndex + 1;
      sheet.insertRowAfter(baseRowIndex);
    }

    // 空行データを作成
    const emptyRow = new Array(numCols).fill('');

    // シートに応じてIDを自動採番
    if (sheetName === SHEETS.PATIENT_MASTER) {
      const idxPid = header.indexOf('patient_id');
      if (idxPid >= 0) {
        emptyRow[idxPid] = generateNextId_(sheet, 'patient_id', 'P', 3);
      }
    } else if (sheetName === SHEETS.STAFF_MASTER) {
      const idxSid = header.indexOf('staff_id');
      if (idxSid >= 0) {
        emptyRow[idxSid] = generateNextId_(sheet, 'staff_id', 'S', 3);
      }
    } else if (sheetName === SHEETS.CHANGE_REQUEST) {
      const idxCid = header.indexOf('change_id');
      if (idxCid >= 0) {
        emptyRow[idxCid] = generateNextId_(sheet, 'change_id', 'C', 3);
      }
    } else if (sheetName === SHEETS.STAFF_CHANGE_REQUEST) {
      const idxScid = header.indexOf('staff_change_id');
      if (idxScid >= 0) {
        emptyRow[idxScid] = generateNextId_(sheet, 'staff_change_id', 'SC', 3);
      }
    } else if (sheetName === SHEETS.EVENT_REQUEST) {
      const idxEvid = header.indexOf('event_id');
      if (idxEvid >= 0) {
        emptyRow[idxEvid] = generateNextId_(sheet, 'event_id', 'EV', 3);
      }
    }

    sheet.getRange(insertAt, 1, 1, numCols).setValues([emptyRow]);

    return {
      success: true,
      message: '行を挿入しました',
      newRowIndex: insertAt,
      newRowData: emptyRow
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 位置情報を更新（選択行のみ）
 * @param {string} sheetName - シート名（患者マスタ or スタッフマスタ）
 * @param {Array<number>} rowIndexes - 更新対象の行番号（1-based）
 * @returns {Object} { success, updatedCount, errors }
 */
function input_updateGeo(sheetName, rowIndexes) {
  requireAdmin_();

  if (!rowIndexes || rowIndexes.length === 0) {
    throw new Error('更新する行を選択してください');
  }

  // 件数制限（Geocoder API制限対策）
  const MAX_ROWS = 20;
  if (rowIndexes.length > MAX_ROWS) {
    throw new Error('一度に更新できるのは' + MAX_ROWS + '件までです。' + rowIndexes.length + '件選択されています。');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('別の処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + sheetName);
    }

    const data = sheet.getDataRange().getValues();
    const header = data[0];

    // 住所列を特定
    let addrHeader = '住所';
    if (sheetName === SHEETS.STAFF_MASTER) {
      addrHeader = '拠点住所';
    }

    const idxAddr = header.indexOf(addrHeader);
    const idxLat = header.indexOf('緯度');
    const idxLng = header.indexOf('経度');

    if (idxAddr < 0 || idxLat < 0 || idxLng < 0) {
      throw new Error('住所/緯度/経度列が見つかりません');
    }

    const geocoder = Maps.newGeocoder();
    let updatedCount = 0;
    const errors = [];

    rowIndexes.forEach(rowIndex => {
      if (rowIndex < 2 || rowIndex > data.length) return;

      const rowData = data[rowIndex - 1]; // 0-based
      const addr = rowData[idxAddr];

      if (!addr) {
        errors.push('行' + rowIndex + ': 住所が空です');
        return;
      }

      try {
        const res = geocoder.geocode(addr);
        if (res.status === 'OK' && res.results && res.results.length > 0) {
          const loc = res.results[0].geometry.location;
          sheet.getRange(rowIndex, idxLat + 1).setValue(loc.lat);
          sheet.getRange(rowIndex, idxLng + 1).setValue(loc.lng);
          updatedCount++;
        } else {
          errors.push('行' + rowIndex + ': 住所「' + addr + '」の位置情報が取得できませんでした');
        }
        Utilities.sleep(200); // API制限対策
      } catch (e) {
        errors.push('行' + rowIndex + ': ' + e.message);
      }
    });

    return {
      success: true,
      message: updatedCount + '件の位置情報を更新しました',
      updatedCount: updatedCount,
      errors: errors
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ID→名前の辞書を取得（自動補完用）
 * @returns {Object} { patients: {id: name}, staff: {id: name} }
 */
function input_getDictionaries() {
  requireAdmin_();

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const result = { patients: {}, staff: {} };

    // 患者マスタ
    const patientSheet = ss.getSheetByName(SHEETS.PATIENT_MASTER);
    if (patientSheet) {
      const pData = patientSheet.getDataRange().getValues();
      if (pData.length > 1) {
        const pHeader = pData[0];
        const idxId = pHeader.indexOf('patient_id');
        const idxName = pHeader.indexOf('患者名');
        if (idxId >= 0 && idxName >= 0) {
          for (let i = 1; i < pData.length; i++) {
            const id = pData[i][idxId];
            const name = pData[i][idxName];
            if (id) result.patients[id] = name || '';
          }
        }
      }
    }

    // スタッフマスタ
    const staffSheet = ss.getSheetByName(SHEETS.STAFF_MASTER);
    if (staffSheet) {
      const sData = staffSheet.getDataRange().getValues();
      if (sData.length > 1) {
        const sHeader = sData[0];
        const idxId = sHeader.indexOf('staff_id');
        const idxName = sHeader.indexOf('スタッフ名');
        if (idxId >= 0 && idxName >= 0) {
          for (let i = 1; i < sData.length; i++) {
            const id = sData[i][idxId];
            const name = sData[i][idxName];
            if (id) result.staff[id] = name || '';
          }
        }
      }
    }

    return { success: true, data: result };
  } catch (e) {
    console.error('input_getDictionaries error:', e);
    return { success: false, error: e.message, data: { patients: {}, staff: {} } };
  }
}

/**
 * エリア候補一覧を取得
 */
function input_getAreaOptions() {
  // 固定候補（実運用ではマスタ化も可）
  return ['A1', 'A2', 'A3', 'B1', 'B2', 'B3'];
}

/**
 * スタッフ選択肢一覧を取得（ID + 名前）
 * @returns {Array<{id: string, name: string, label: string}>}
 */
function input_getStaffOptions() {
  requireAdmin_();

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.STAFF_MASTER);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const header = data[0];
    const idxId = header.indexOf('staff_id');
    const idxName = header.indexOf('スタッフ名');
    if (idxId < 0 || idxName < 0) return [];

    const options = [];
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][idxId] || '').trim();
      const name = String(data[i][idxName] || '').trim();
      if (id) {
        options.push({
          id: id,
          name: name,
          label: id + ' ' + name
        });
      }
    }
    return options;
  } catch (e) {
    console.error('input_getStaffOptions error:', e);
    return [];
  }
}

/**
 * 次のIDを生成（P001, S001, C001形式）
 * @param {Sheet} sheet - 対象シート
 * @param {string} idHeaderName - IDのヘッダー名（patient_id, staff_id, change_id）
 * @param {string} prefix - プレフィックス（P, S, C）
 * @param {number} padLen - ゼロ埋め桁数（デフォルト3）
 * @returns {string} 新しいID
 */
function generateNextId_(sheet, idHeaderName, prefix, padLen) {
  padLen = padLen || 3;

  if (!sheet) return prefix + '001';

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return prefix + '001';

  const header = data[0];
  const idxId = header.indexOf(idHeaderName);
  if (idxId < 0) return prefix + '001';

  let maxNum = 0;
  const regex = new RegExp('^' + prefix + '(\\d+)$', 'i');

  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][idxId] || '').trim();
    const match = id.match(regex);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }

  const nextNum = maxNum + 1;
  return prefix + String(nextNum).padStart(padLen, '0');
}

/**
 * ウィザードから行を作成
 * @param {string} formType - フォームタイプ（患者マスタ, スタッフマスタ, 個別変更リクエスト）
 * @param {Object} answers - 回答オブジェクト { key: value, ... }
 * @param {number} insertAfterRow - 挿入位置（1-based、省略時は末尾）
 * @returns {Object} { success, message, newRowIndex, newRowData }
 */
function input_createRowFromWizard(formType, answers, insertAfterRow) {
  var lock = null;
  try {
    // 権限チェック
    requireAdmin_();

    lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return { success: false, error: '別の処理が実行中です。少し待ってから再実行してください。' };
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // フォームタイプからシート名を決定
    var sheetName;
    if (formType === '患者マスタ') {
      sheetName = SHEETS.PATIENT_MASTER;
    } else if (formType === 'スタッフマスタ') {
      sheetName = SHEETS.STAFF_MASTER;
    } else if (formType === '個別変更リクエスト') {
      sheetName = SHEETS.CHANGE_REQUEST;
    } else if (formType === 'スタッフ個別変更') {
      sheetName = SHEETS.STAFF_CHANGE_REQUEST;
    } else {
      return { success: false, error: '不明なフォームタイプ: ' + formType };
    }

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, error: 'シートが見つかりません: ' + sheetName };
    }

    var numCols = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    var header = sheet.getRange(1, 1, 1, numCols).getValues()[0];

    // 挿入位置を決定
    var insertAt;
    if (!insertAfterRow || insertAfterRow < 2) {
      insertAt = lastRow + 1;
    } else if (insertAfterRow >= lastRow) {
      insertAt = lastRow + 1;
    } else {
      insertAt = insertAfterRow + 1;
      sheet.insertRowAfter(insertAfterRow);
    }

    // 空行データを作成
    var rowData = new Array(numCols).fill('');

    // IDを自動採番（autoIdタイプの場合）
    if (formType === '患者マスタ') {
      var idxPid = header.indexOf('patient_id');
      if (idxPid >= 0) {
        rowData[idxPid] = generateNextId_(sheet, 'patient_id', 'P', 3);
      }
    } else if (formType === 'スタッフマスタ') {
      var idxSid = header.indexOf('staff_id');
      if (idxSid >= 0) {
        rowData[idxSid] = generateNextId_(sheet, 'staff_id', 'S', 3);
      }
    } else if (formType === '個別変更リクエスト') {
      var idxCid = header.indexOf('change_id');
      if (idxCid >= 0) {
        rowData[idxCid] = generateNextId_(sheet, 'change_id', 'C', 3);
      }
    } else if (formType === 'スタッフ個別変更') {
      var idxScid = header.indexOf('staff_change_id');
      if (idxScid >= 0) {
        rowData[idxScid] = generateNextId_(sheet, 'staff_change_id', 'SC', 3);
      }
    } else if (formType === 'イベントリクエスト') {
      var idxEvid = header.indexOf('event_id');
      if (idxEvid >= 0) {
        rowData[idxEvid] = generateNextId_(sheet, 'event_id', 'EV', 3);
      }
    }

    // ヘッダー名とキーのマッピング定義
    var headerMapping = {
      // 患者マスタ用
      'name': '患者名',
      'sex': '性別',
      'address': '住所',
      'lat': '緯度',
      'lng': '経度',
      'area': 'エリア',
      'weeklyCount': '週訪問回数',
      'preferDays': '希望曜日（複数可）',
      'ngDays': '曜日NG',
      'timeType': '時間タイプ',
      'timeStart': '希望時間帯（開始）',
      'timeEnd': '希望時間帯（終了）',
      'serviceMin': 'サービス時間',
      'sexLimit': '性別制限',
      'needStaff': '必要スタッフ数',
      'fixedStaff': '指定スタッフID',
      'staffType': '指定タイプ',
      'ngStaff': 'NGスタッフID',
      'contPref': '継続希望',
      'note': '備考',
      // スタッフマスタ用
      'staffName': 'スタッフ名',
      'baseAddress': '拠点住所',
      'shiftStart': 'シフト開始',
      'shiftEnd': 'シフト終了',
      'workDays': '勤務曜日',
      'areas': '得意エリア',
      'maxPerDay': '最大訪問件数/日',
      'skill': 'スキル',
      // 個別変更リクエスト用
      'date': '日付',
      'operation': '操作',
      'start': '新開始時刻',
      'end': '新終了時刻',
      // スタッフ個別変更リクエスト用
      'restrictionType': '制限タイプ',
      'startTime': '開始時刻',
      'endTime': '終了時刻',
      'reason': '理由',
      // イベントリクエスト用
      'eventType': 'イベント種別',
      'title': 'タイトル',
      'timeMode': '時間指定方法',
      'durationMin': '所要時間(分)',
      'fixedSlot': '固定枠',
      'patientLinked': '患者紐づき',
      'patientAffect': '患者影響',
      'returnBefore': '事前事務所戻り',
      'returnAfter': '事後事務所戻り',
      'notes': '備考'
    };

    // 曜日の日本語→英語変換マップ
    var youbiJpToEn = {
      '日': 'Sun', '月': 'Mon', '火': 'Tue', '水': 'Wed',
      '木': 'Thu', '金': 'Fri', '土': 'Sat'
    };

    // 自動生成されたIDフィールドのリスト（上書き禁止）
    var autoIdFields = ['patient_id', 'staff_id', 'change_id', 'staff_change_id', 'event_id'];

    // answersをrowDataにマッピング
    for (var key in answers) {
      if (!answers.hasOwnProperty(key)) continue;
      var value = answers[key];

      // 空値はスキップ
      if (value === undefined || value === null || value === '') continue;

      // ヘッダー名を取得
      var headerName = headerMapping[key] || key;
      var idx = header.indexOf(headerName);
      if (idx < 0) continue;

      // 自動生成IDフィールドで既に値がある場合はスキップ（上書き防止）
      if (autoIdFields.indexOf(headerName) >= 0 && rowData[idx]) continue;

      // 配列の場合（multiSelect）はCSVに変換
      if (Array.isArray(value)) {
        // 曜日フィールドの場合は日本語→英語変換
        if (key === 'preferDays' || key === 'workDays') {
          value = value.map(function(d) {
            return youbiJpToEn[d] || d;
          });
        }
        rowData[idx] = value.join(',');
      } else {
        // 曜日NGフィールドの場合は日本語→英語変換
        if (key === 'ngDays' && value) {
          value = youbiJpToEn[value] || value;
        }
        rowData[idx] = value;
      }
    }

    // 個別変更リクエストの場合、患者IDから患者名を自動取得
    if (formType === '個別変更リクエスト') {
      var idxPatientId = header.indexOf('patient_id');
      var idxPatientName = header.indexOf('患者名');
      if (idxPatientId >= 0 && idxPatientName >= 0 && rowData[idxPatientId] && !rowData[idxPatientName]) {
        var patientSheet = ss.getSheetByName(SHEETS.PATIENT_MASTER);
        if (patientSheet) {
          var patientData = patientSheet.getDataRange().getValues();
          if (patientData.length > 1) {
            var pHeader = patientData[0];
            var pIdIdx = pHeader.indexOf('patient_id');
            var pNameIdx = pHeader.indexOf('患者名');
            if (pIdIdx >= 0 && pNameIdx >= 0) {
              var targetPid = rowData[idxPatientId];
              for (var p = 1; p < patientData.length; p++) {
                if (String(patientData[p][pIdIdx]).trim() === String(targetPid).trim()) {
                  rowData[idxPatientName] = patientData[p][pNameIdx] || '';
                  break;
                }
              }
            }
          }
        }
      }
    }

    // スタッフ個別変更の場合、スタッフIDからスタッフ名を自動取得 + 曜日を自動計算
    if (formType === 'スタッフ個別変更') {
      // スタッフ名の自動取得
      var idxStaffId = header.indexOf('staff_id');
      var idxStaffName = header.indexOf('スタッフ名');
      if (idxStaffId >= 0 && idxStaffName >= 0 && rowData[idxStaffId] && !rowData[idxStaffName]) {
        var staffSheet = ss.getSheetByName(SHEETS.STAFF_MASTER);
        if (staffSheet) {
          var staffData = staffSheet.getDataRange().getValues();
          if (staffData.length > 1) {
            var sHeader = staffData[0];
            var sIdIdx = sHeader.indexOf('staff_id');
            var sNameIdx = sHeader.indexOf('スタッフ名');
            if (sIdIdx >= 0 && sNameIdx >= 0) {
              var targetSid = rowData[idxStaffId];
              for (var s = 1; s < staffData.length; s++) {
                if (String(staffData[s][sIdIdx]).trim() === String(targetSid).trim()) {
                  rowData[idxStaffName] = staffData[s][sNameIdx] || '';
                  break;
                }
              }
            }
          }
        }
      }

      // 曜日の自動計算（日付から）
      var idxDate = header.indexOf('日付');
      var idxYoubi = header.indexOf('曜日');
      if (idxDate >= 0 && idxYoubi >= 0 && rowData[idxDate] && !rowData[idxYoubi]) {
        var dateVal = rowData[idxDate];
        var dateObj;
        if (dateVal instanceof Date) {
          dateObj = dateVal;
        } else {
          dateObj = new Date(dateVal);
        }
        if (!isNaN(dateObj.getTime())) {
          var youbiNames = ['日', '月', '火', '水', '木', '金', '土'];
          rowData[idxYoubi] = youbiNames[dateObj.getDay()];
        }
      }
    }

    // 行を挿入
    sheet.getRange(insertAt, 1, 1, numCols).setValues([rowData]);

    // 生成されたIDを取得
    var generatedId = '';
    if (formType === '患者マスタ') {
      generatedId = rowData[header.indexOf('patient_id')] || '';
    } else if (formType === 'スタッフマスタ') {
      generatedId = rowData[header.indexOf('staff_id')] || '';
    } else if (formType === '個別変更リクエスト') {
      generatedId = rowData[header.indexOf('change_id')] || '';
    } else if (formType === 'スタッフ個別変更') {
      generatedId = rowData[header.indexOf('staff_change_id')] || '';
    }

    return {
      success: true,
      message: formType + 'に新しいレコードを追加しました（ID: ' + generatedId + '）',
      newRowIndex: insertAt,
      newRowData: rowData,
      generatedId: generatedId
    };
  } catch (e) {
    console.error('input_createRowFromWizard error:', e);
    return { success: false, error: e.message || String(e) };
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (ignore) {}
    }
  }
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

// ============================================================
// 共通ユーティリティ関数
// ============================================================

function normalizeYoubi(y) {
  if (!y) return null;
  y = String(y).trim();

  var en = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  if (en.indexOf(y) >= 0) return y;

  var enFull = {
    'Sunday':'Sun','Monday':'Mon','Tuesday':'Tue','Wednesday':'Wed',
    'Thursday':'Thu','Friday':'Fri','Saturday':'Sat'
  };
  if (enFull[y]) return enFull[y];

  var jp1 = { '日':'Sun','月':'Mon','火':'Tue','水':'Wed','木':'Thu','金':'Fri','土':'Sat' };
  if (jp1[y]) return jp1[y];

  var jp2 = { '日曜':'Sun','月曜':'Mon','火曜':'Tue','水曜':'Wed','木曜':'Thu','金曜':'Fri','土曜':'Sat' };
  if (jp2[y]) return jp2[y];

  var jp3 = { '日曜日':'Sun','月曜日':'Mon','火曜日':'Tue','水曜日':'Wed','木曜日':'Thu','金曜日':'Fri','土曜日':'Sat' };
  if (jp3[y]) return jp3[y];

  return null;
}

function toHalfWidthNumber_(v, def) {
  if (v === null || v === undefined || v === '') return def;
  var s = String(v).trim().replace(/[０-９]/g, function(c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
  var n = parseInt(s, 10);
  return isNaN(n) ? def : n;
}

function parseIdList(str) {
  if (!str) return [];
  return String(str).split(/[,\u3001\/・\s]+/).map(function(s){ return s.trim(); }).filter(Boolean);
}

function parseTypeList(str) {
  if (!str) return [];
  return String(str).split(/[,\u3001\/・\s]+/).map(function(s){ return s.trim(); }).filter(Boolean);
}

// ============================================================
// 週ビューを更新（ss引数版）
// ============================================================

function 週ビューを更新_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const resultSheet  = ss.getSheetByName('割当結果');
  const viewSheet    = ss.getSheetByName('週ビュー');
  const staffSheet   = ss.getSheetByName('スタッフマスタ');
  const patientSheet = ss.getSheetByName('患者マスタ');

  if (!resultSheet || !viewSheet) {
    throw new Error('「割当結果」シートと「週ビュー」シートを作ってから実行してください。');
  }

  const staffGenderMap = {};
  if (staffSheet) {
    const sValues = staffSheet.getDataRange().getValues();
    if (sValues.length > 1) {
      const sHeader = sValues[0];
      const sData   = sValues.slice(1);
      const sIdxId   = sHeader.indexOf('staff_id');
      const sIdxName = sHeader.indexOf('スタッフ名');
      const sIdxGen  = sHeader.indexOf('性別');
      sData.forEach(r => {
        const id = r[sIdxId];
        if (!id) return;
        staffGenderMap[id] = {
          gender: sIdxGen >= 0 ? (r[sIdxGen] || '') : '',
          name  : sIdxName >= 0 ? (r[sIdxName] || '') : ''
        };
      });
    }
  }

  const patientGenderMap = {};
  if (patientSheet) {
    const pValues = patientSheet.getDataRange().getValues();
    if (pValues.length > 1) {
      const pHeader = pValues[0];
      const pData   = pValues.slice(1);
      const pIdxId  = pHeader.indexOf('patient_id');
      const pIdxGen = pHeader.indexOf('性別');
      const pIdxName = pHeader.indexOf('患者名');
      pData.forEach(r => {
        const id = r[pIdxId];
        if (!id) return;
        patientGenderMap[id] = {
          gender: pIdxGen >= 0 ? (r[pIdxGen] || '') : '',
          name: pIdxName >= 0 ? (r[pIdxName] || '') : ''
        };
      });
    }
  }

  // ※イベントは割当結果のEV行から取得するため、イベントリクエストの読み込みは不要

  const today = new Date();
  today.setHours(0,0,0,0);
  const day = today.getDay();
  const diffToMonday = (day + 6) % 7;
  const start = new Date(today);
  start.setDate(today.getDate() - diffToMonday);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);

  const values = resultSheet.getDataRange().getValues();
  if (values.length <= 1) {
    throw new Error('「割当結果」にデータがありません。');
  }
  const header = values[0];
  const data   = values.slice(1);

  const idxDate    = header.indexOf('日付');
  const idxStaff   = header.indexOf('スタッフ名');
  const idxStaffId = header.indexOf('staff_id');
  const idxStart   = header.indexOf('開始時刻');
  const idxEnd     = header.indexOf('終了時刻');
  const idxPatient = header.indexOf('患者名');
  const idxPid     = header.indexOf('patient_id');
  const idxVisitId = header.indexOf('visit_id');
  const idxNote    = header.indexOf('備考');

  if ([idxDate,idxStaff,idxStaffId,idxStart,idxEnd,idxPatient,idxPid].some(i => i === -1)) {
    throw new Error('「割当結果」のヘッダー名を確認してください。');
  }

  const startStr = Utilities.formatDate(start, tz, 'yyyy/MM/dd');
  const endStr   = Utilities.formatDate(end,   tz, 'yyyy/MM/dd');

  const weekData = data.filter(row => {
    const d = row[idxDate];
    if (!(d instanceof Date)) return false;
    const ds = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    return ds >= startStr && ds <= endStr;
  });

  const staffMap = new Map();
  // 訪問からスタッフを収集
  weekData.forEach(r => {
    const sid   = r[idxStaffId];
    const sname = r[idxStaff] || '';
    if (!sid && !sname) return;
    const key = sid || sname;
    if (!staffMap.has(key)) {
      let gender = '';
      if (sid && staffGenderMap[sid]) gender = staffGenderMap[sid].gender || '';
      staffMap.set(key, { id: sid || '', name: sname, gender: gender });
    }
  });

  // ※イベントは割当結果のEV行に含まれるため、staffMapには既に入っている

  let staffList = Array.from(staffMap.values());
  staffList.sort((a,b) => {
    if (a.name === '未割当' && b.name !== '未割当') return 1;
    if (b.name === '未割当' && a.name !== '未割当') return -1;
    return a.name.localeCompare(b.name, 'ja');
  });

  viewSheet.clear();
  viewSheet.getRange(1,1).setValue('職員名');

  for (let i = 0; i < 7; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const label = Utilities.formatDate(d, tz, 'MM/dd(EEE)');
    viewSheet.getRange(1, 2 + i).setValue(label);
  }

  staffList.forEach((st, idx) => {
    let label = '';
    if (st.id) label += st.id + ' ';
    label += st.name || '';
    if (st.gender) label += '（' + st.gender + '）';
    viewSheet.getRange(2 + idx, 1).setValue(label);
  });

  // 時間値をHH:mm形式に変換するヘルパー関数
  function formatTimeVal(val) {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, tz, 'HH:mm');
    if (typeof val === 'number') {
      const base = new Date(1899, 11, 30);
      const ms   = val * 24 * 60 * 60 * 1000;
      const dd   = new Date(base.getTime() + ms);
      return Utilities.formatDate(dd, tz, 'HH:mm');
    }
    return String(val);
  }

  // 時間値を分に変換するヘルパー関数（ソート用）
  function toSortMinutes(val) {
    if (!val) return 9999;
    if (val instanceof Date) return val.getHours() * 60 + val.getMinutes();
    if (typeof val === 'number') {
      // シリアル値から分に変換
      const totalMinutes = Math.round(val * 24 * 60);
      return totalMinutes % 1440;
    }
    // "HH:mm" 形式の文字列
    const match = String(val).match(/(\d{1,2}):(\d{2})/);
    if (match) return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
    return 9999;
  }

  // "その行がイベントか" 判定（visit_id が EV_ か、備考に [EV]）
  function isEventRowFromResult_(row) {
    const vid  = (idxVisitId >= 0) ? String(row[idxVisitId] || '') : '';
    const note = (idxNote >= 0) ? String(row[idxNote] || '') : '';
    return vid.indexOf('EV_') === 0 || note.indexOf('[EV]') >= 0;
  }

  let eventCount = 0;

  staffList.forEach((st, rIndex) => {
    for (let i = 0; i < 7; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      const targetDateStr = Utilities.formatDate(d, tz, 'yyyy/MM/dd');

      // 割当結果から、そのスタッフ×その日の行だけ取る（EV行もここに入る）
      const rows = weekData.filter(row => {
        const d2 = row[idxDate];
        const ds2 = Utilities.formatDate(d2, tz, 'yyyy/MM/dd');
        const sid   = row[idxStaffId] || '';
        const sname = row[idxStaff]   || '';
        const key   = sid || sname;
        return key === (st.id || st.name) && ds2 === targetDateStr;
      });

      const displayItems = [];

      rows.forEach(v => {
        const startVal = v[idxStart];
        const endVal   = v[idxEnd];
        const pid      = v[idxPid] || '';
        const pname    = v[idxPatient] || '';
        const noteVal  = (idxNote >= 0) ? String(v[idxNote] || '') : '';
        const vid      = (idxVisitId >= 0) ? String(v[idxVisitId] || '') : '';

        const stime = formatTimeVal(startVal);
        const etime = formatTimeVal(endVal);

        // EV行として表示（割当結果の備考に [EV] が入っている）
        if (isEventRowFromResult_(v)) {
          eventCount++;
          let evText = '';
          if (stime && etime) evText += stime + '〜' + etime + ' ';
          evText += noteVal ? noteVal : '[EV]';

          displayItems.push({
            sortKey: toSortMinutes(startVal),
            text: evText,
            isEvent: true
          });
          return;
        }

        // 通常訪問として表示
        const pInfo   = pid ? patientGenderMap[pid] : null;
        const pGender = pInfo ? (pInfo.gender || '') : '';
        const isTwo   = (vid.indexOf('-') >= 0) || (noteVal.indexOf('同時訪問') >= 0);
        const mark    = isTwo ? '👥 ' : '';

        let pidPart = '';
        if (pid) {
          pidPart = pid;
          if (pGender) pidPart += '（' + pGender + '）';
          pidPart += ' ';
        }

        let text;
        if (!stime && !etime) {
          text = mark + pidPart + pname;
        } else {
          text = mark + stime + '〜' + etime + ' ' + pidPart + pname;
        }

        displayItems.push({
          sortKey: toSortMinutes(startVal),
          text: text,
          isEvent: false
        });
      });

      // 時刻でソート
      displayItems.sort((a, b) => a.sortKey - b.sortKey);

      const cellText = displayItems.map(item => item.text).join('\n');
      if (cellText) {
        const cell = viewSheet.getRange(2 + rIndex, 2 + i);
        cell.setValue(cellText);
        cell.setWrap(true);
      }
    }
  });

  return { message: '週ビューを更新しました（' + staffList.length + '名、EV行' + eventCount + '件含む）' };
}

// ============================================================
// 割当結果を作成（ss引数版）
// ============================================================

function 割当結果を作成_(ss) {
  const tz = ss.getSpreadsheetTimeZone();

  const weeklySheet = ss.getSheetByName('週間リクエスト');
  const staffSheet  = ss.getSheetByName('スタッフマスタ');
  const resultSheet = ss.getSheetByName('割当結果');
  let historySheet = ss.getSheetByName('訪問履歴');
  const patientSheet = ss.getSheetByName('患者マスタ');

  if (!patientSheet) throw new Error('「患者マスタ」シートがありません。');
  if (!weeklySheet || !staffSheet || !resultSheet) {
    throw new Error('「週間リクエスト」「スタッフマスタ」「割当結果」シートがあるか確認してください。');
  }

  if (!historySheet) historySheet = ss.insertSheet('訪問履歴');
  if (historySheet.getLastRow() === 0) {
    var histHeader = ['visit_id','日付','曜日','staff_id','スタッフ名','patient_id','患者名','エリア','開始時刻','終了時刻','サービス時間','備考'];
    historySheet.getRange(1, 1, 1, histHeader.length).setValues([histHeader]);
  }

  function calcDistanceKm(lat1, lng1, lat2, lng2) {
    if (lat1 == null || lng1 == null || lat2 == null || lng2 == null) return null;
    const R = 6371;
    const toRad = Math.PI / 180;
    const dLat = (lat2 - lat1) * toRad;
    const dLng = (lng2 - lng1) * toRad;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * toRad) * Math.cos(lat2 * toRad) *
              Math.sin(dLng/2) * Math.sin(dLng/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
  }

  function distToScore(km) {
    if (km == null) return 99;
    if (km <= 2)  return 0;
    if (km <= 5)  return 1;
    if (km <= 10) return 2;
    return 5;
  }

  function applyStaffPreference(candidates, specifiedIdsArr, specifiedType, ngIdsArr) {
    var ngSet = {};
    ngIdsArr.forEach(function(id){ ngSet[id] = true; });
    candidates = candidates.filter(function(c){ return !ngSet[c.staff.id]; });
    if (!specifiedType || specifiedIdsArr.length === 0) return candidates;
    var specSet = {};
    specifiedIdsArr.forEach(function(id){ specSet[id] = true; });
    if (specifiedType === '必須') return candidates.filter(function(c){ return specSet[c.staff.id]; });
    if (specifiedType === '優先') {
      candidates.forEach(function(c){ c._pref = specSet[c.staff.id] ? 1 : 0; });
      candidates.sort(function(a,b){ if (a._pref !== b._pref) return b._pref - a._pref; return 0; });
    }
    return candidates;
  }

  var wValues = weeklySheet.getDataRange().getValues();
  if (wValues.length <= 1) throw new Error('「週間リクエスト」にデータがありません。');
  var wHeader = wValues[0];
  var wData   = wValues.slice(1);
  var wIdx = {
    date: wHeader.indexOf('日付'), youbi: wHeader.indexOf('曜日'), pid: wHeader.indexOf('patient_id'),
    pname: wHeader.indexOf('患者名'), area: wHeader.indexOf('エリア'), start: wHeader.indexOf('開始時刻'),
    end: wHeader.indexOf('終了時刻'), svcMin: wHeader.indexOf('サービス時間'), needStaff: wHeader.indexOf('必要スタッフ数'),
    specifiedIds: wHeader.indexOf('指定スタッフID'), specifiedType: wHeader.indexOf('指定タイプ'),
    ngStaffIds: wHeader.indexOf('NGスタッフID'), sexLimit: wHeader.indexOf('性別制限'),
    contPref: wHeader.indexOf('継続希望'), change: wHeader.indexOf('変更区分（通常/変更/追加/キャンセル）'),
    prevSid: wHeader.indexOf('前回担当スタッフID'), prevSname: wHeader.indexOf('前回担当スタッフ名'),
    timeType: wHeader.indexOf('時間タイプ'), earliest: wHeader.indexOf('希望最早時刻'),
    latest: wHeader.indexOf('希望最遅時刻'), note: wHeader.indexOf('備考')
  };

  var pValues = patientSheet.getDataRange().getValues();
  var pHeader = pValues[0];
  var pData   = pValues.slice(1);
  var pIdx = { id: pHeader.indexOf('patient_id'), name: pHeader.indexOf('患者名'), area: pHeader.indexOf('エリア'),
               lat: pHeader.indexOf('緯度'), lng: pHeader.indexOf('経度'), svcMin: pHeader.indexOf('サービス時間') };

  var patientMap = {};
  pData.forEach(function(row){
    var id = row[pIdx.id];
    if (!id) return;
    patientMap[id] = { name: row[pIdx.name], area: row[pIdx.area], lat: Number(row[pIdx.lat]) || null,
                       lng: Number(row[pIdx.lng]) || null, svcMin: Number(row[pIdx.svcMin]) || 0 };
  });

  var sValues = staffSheet.getDataRange().getValues();
  if (sValues.length <= 1) throw new Error('「スタッフマスタ」にスタッフが1人もいません。');
  var sHeader = sValues[0];
  var sData   = sValues.slice(1);
  var sIdx = { id: sHeader.indexOf('staff_id'), name: sHeader.indexOf('スタッフ名'), gender: sHeader.indexOf('性別'),
               lat: sHeader.indexOf('緯度'), lng: sHeader.indexOf('経度'), shiftS: sHeader.indexOf('シフト開始'),
               shiftE: sHeader.indexOf('シフト終了'), days: sHeader.indexOf('勤務曜日'), areas: sHeader.indexOf('得意エリア'),
               maxPer: sHeader.indexOf('最大訪問件数/日') };

  function parseDays(str) {
    if (!str) return [];
    var parts = String(str).split(/[,\u3001\/・\s]+/);
    var out = [];
    parts.forEach(function(p){ p = p.trim(); if (!p) return; var y = normalizeYoubi(p); if (y && out.indexOf(y) === -1) out.push(y); });
    return out;
  }

  // 文字列日付もDate化するパーサ
  function parseDate_(v) {
    if (!v) return null;
    if (v instanceof Date) return v;

    var s = String(v).trim();
    var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
      var dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      return isNaN(dt.getTime()) ? null : dt;
    }
    var dt2 = new Date(s);
    return isNaN(dt2.getTime()) ? null : dt2;
  }

  // "10:00" / シリアル / Date すべて分にするパーサ
  function parseTimeToMinutes_(v) {
    if (v === null || v === undefined || v === '') return null;

    // 数値（シリアル値）の場合
    if (typeof v === 'number') {
      // 0〜1の範囲ならシリアル時刻として扱う
      if (v >= 0 && v < 1) {
        return Math.round(v * 24 * 60);
      }
      // 1以上なら分として解釈（誤入力対応）
      if (v >= 1 && v < 1440) {
        return Math.round(v);
      }
      // それ以外はシリアル値として扱う
      return Math.round(v * 24 * 60) % 1440;
    }

    // Date オブジェクトの場合
    if (v instanceof Date) {
      return v.getHours() * 60 + v.getMinutes();
    }

    // 文字列の場合
    var s = String(v).trim();

    // "14:00" や "9:30" 形式
    var m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (m) return Number(m[1]) * 60 + Number(m[2]);

    // "14:00:00" 形式（秒付き）
    m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
    if (m) return Number(m[1]) * 60 + Number(m[2]);

    // "午前10時30分" や "10時30分" 形式
    m = s.match(/(\d{1,2})時(\d{1,2})分/);
    if (m) return Number(m[1]) * 60 + Number(m[2]);

    // "14時" 形式（分なし）
    m = s.match(/^(\d{1,2})時$/);
    if (m) return Number(m[1]) * 60;

    // "HH:mm〜HH:mm" などにも保険で対応（先頭だけ）
    m = s.match(/(\d{1,2}):(\d{2})/);
    if (m) return Number(m[1]) * 60 + Number(m[2]);

    return null;
  }

  // 分→シリアル(スプレッドシート時刻) に変換
  function minutesToSerial_(min) {
    if (min == null) return '';
    return min / (24 * 60);
  }

  function toMinutes(v) {
    return parseTimeToMinutes_(v);
  }

  function normalizeContPref(v) {
    if (!v) return '';
    v = String(v).trim();
    if (v === '同じ人' || v === '同じ人希望') return '同じ人希望';
    if (v === 'ローテーション優先') return 'ローテーション優先';
    if (v === 'どちらでも') return 'どちらでも';
    return v;
  }

  var EXTRA_BUFFER_MIN = 15;

  var staffList = [];
  sData.forEach(function(row){
    var id = row[sIdx.id];
    var name = row[sIdx.name];
    if (!id || !name) return;
    var workDays = parseDays(row[sIdx.days]);
    var shiftStartMin = toMinutes(row[sIdx.shiftS]);
    var shiftEndMin   = toMinutes(row[sIdx.shiftE]);
    var areasStr = row[sIdx.areas] || '';
    var areaList = String(areasStr).split(/[,\u3001\/・\s]+/).map(function(s){ return s.trim(); }).filter(function(s){ return s; });
    var maxPerDay = Number(row[sIdx.maxPer] || 0) || 999;
    staffList.push({ id: id, name: name, gender: row[sIdx.gender] || '', lat: row[sIdx.lat], lng: row[sIdx.lng],
                     shiftStartMin: shiftStartMin, shiftEndMin: shiftEndMin, workDays: workDays, areas: areaList, maxPerDay: maxPerDay });
  });

  if (staffList.length === 0) throw new Error('有効なスタッフ情報がありません。');

  // ============================================================
  // Task A: スタッフ個別変更リクエストの読み込みとMap化
  // ============================================================
  var staffChangeMap = {};  // key: staff_id|yyyy/MM/dd => [records...]
  var staffChangeSheet = ss.getSheetByName('スタッフ個別変更リクエスト');
  if (staffChangeSheet && staffChangeSheet.getLastRow() > 1) {
    var scValues = staffChangeSheet.getDataRange().getValues();
    var scHeader = scValues[0];
    var scIdx = {
      staffId: scHeader.indexOf('staff_id'),
      date: scHeader.indexOf('日付'),
      restrictionType: scHeader.indexOf('制限タイプ'),
      startTime: scHeader.indexOf('開始時刻'),
      endTime: scHeader.indexOf('終了時刻')
    };

    for (var sci = 1; sci < scValues.length; sci++) {
      var scRow = scValues[sci];
      var scStaffId = scRow[scIdx.staffId];
      // parseDate_ で文字列日付もパース
      var scDateObj = parseDate_(scRow[scIdx.date]);
      if (!scStaffId || !scDateObj) continue;

      var scDateStr = Utilities.formatDate(scDateObj, tz, 'yyyy/MM/dd');

      var scKey = scStaffId + '|' + scDateStr;
      if (!staffChangeMap[scKey]) staffChangeMap[scKey] = [];

      staffChangeMap[scKey].push({
        restrictionType: String(scRow[scIdx.restrictionType] || '').trim(),
        startTime: toMinutes(scRow[scIdx.startTime]),
        endTime: toMinutes(scRow[scIdx.endTime])
      });
    }
  }

  // ============================================================
  // Task A-2: イベントリクエストの読み込みとMap化
  // ============================================================
  var eventMap = {};  // key: staff_id|yyyy/MM/dd => [events...]
  var eventSheet = ss.getSheetByName('イベントリクエスト');
  if (eventSheet && eventSheet.getLastRow() > 1) {
    var evValues = eventSheet.getDataRange().getValues();
    var evHeader = evValues[0];

    // ヘッダー検索ヘルパー（空白や文字の違いに対応）
    function findEvHeaderIndex(headers, targetName) {
      var idx = headers.indexOf(targetName);
      if (idx >= 0) return idx;
      for (var i = 0; i < headers.length; i++) {
        if (String(headers[i]).trim() === targetName) return i;
      }
      for (var j = 0; j < headers.length; j++) {
        var h = String(headers[j]).trim();
        if (h.indexOf(targetName) >= 0 || targetName.indexOf(h) >= 0) return j;
      }
      return -1;
    }

    var evIdx = {
      eventId: findEvHeaderIndex(evHeader, 'event_id'),
      staffId: findEvHeaderIndex(evHeader, 'staff_id'),
      date: findEvHeaderIndex(evHeader, '日付'),
      eventType: findEvHeaderIndex(evHeader, 'イベント種別'),
      title: findEvHeaderIndex(evHeader, 'タイトル'),
      timeMode: findEvHeaderIndex(evHeader, '時間指定方法'),
      startTime: findEvHeaderIndex(evHeader, '開始時刻'),
      endTime: findEvHeaderIndex(evHeader, '終了時刻'),
      durationMin: findEvHeaderIndex(evHeader, '所要時間'),
      fixedSlot: findEvHeaderIndex(evHeader, '固定枠'),
      patientLinked: findEvHeaderIndex(evHeader, '患者紐づき'),
      patientId: findEvHeaderIndex(evHeader, 'patient_id'),
      patientAffect: findEvHeaderIndex(evHeader, '患者影響'),
      lat: findEvHeaderIndex(evHeader, '緯度'),
      lng: findEvHeaderIndex(evHeader, '経度')
    };

    for (var evi = 1; evi < evValues.length; evi++) {
      var evRow = evValues[evi];
      var evStaffId = evRow[evIdx.staffId];
      // parseDate_ で文字列日付もパース
      var evDateObj = parseDate_(evRow[evIdx.date]);
      if (!evStaffId || !evDateObj) continue;
      var evDateStr = Utilities.formatDate(evDateObj, tz, 'yyyy/MM/dd');

      var evKey = evStaffId + '|' + evDateStr;
      if (!eventMap[evKey]) eventMap[evKey] = [];

      // 時間をminに変換（toMinutes は parseTimeToMinutes_ をラップ）
      var evStartMin = toMinutes(evRow[evIdx.startTime]);
      var evEndMin = toMinutes(evRow[evIdx.endTime]);
      // duration: 値があれば使用、なければ開始/終了から計算、それもなければ60分
      var rawEvDuration = evRow[evIdx.durationMin];
      var evDuration;
      if (rawEvDuration && !isNaN(parseInt(rawEvDuration, 10))) {
        evDuration = parseInt(rawEvDuration, 10);
      } else if (evStartMin != null && evEndMin != null && evEndMin > evStartMin) {
        evDuration = evEndMin - evStartMin;  // 開始/終了から自動算出
      } else {
        evDuration = 60;  // デフォルト
      }
      var evTimeMode = String(evRow[evIdx.timeMode] || '').trim();

      eventMap[evKey].push({
        eventId: evRow[evIdx.eventId] || '',
        eventType: evRow[evIdx.eventType] || '',
        title: evRow[evIdx.title] || '',
        dateObj: evDateObj,
        timeMode: evTimeMode,
        startMin: evStartMin,
        endMin: evEndMin,
        durationMin: evDuration,
        fixedSlot: evRow[evIdx.fixedSlot] === true || evRow[evIdx.fixedSlot] === 'TRUE',
        patientLinked: evRow[evIdx.patientLinked] === true || evRow[evIdx.patientLinked] === 'TRUE',
        patientId: evRow[evIdx.patientId] || null,
        patientAffect: String(evRow[evIdx.patientAffect] || '').trim(),
        lat: evRow[evIdx.lat],
        lng: evRow[evIdx.lng]
      });
    }
  }

  // ============================================================
  // Task B & C: スタッフ制限の不可区間取得と衝突判定
  // ============================================================

  // スタッフの基本シフト情報を取得するヘルパー
  function getStaffShift_(staffId) {
    for (var i = 0; i < staffList.length; i++) {
      if (staffList[i].id === staffId) {
        return { shiftStartMin: staffList[i].shiftStartMin, shiftEndMin: staffList[i].shiftEndMin };
      }
    }
    return { shiftStartMin: 0, shiftEndMin: 1440 };
  }

  // スタッフの不可区間を取得（制限タイプ＋イベントに基づいて正規化）
  function getStaffBlockedIntervals_(staffId, dateStr) {
    var shift = getStaffShift_(staffId);
    var intervals = [];

    // スタッフ個別変更リクエストからの制限
    var records = staffChangeMap[staffId + '|' + dateStr];
    if (records && records.length > 0) {
      records.forEach(function(rec) {
        var rType = rec.restrictionType;

        if (rType === '休み' || rType === '終日不可' || rType === '終日') {
          // 終日不可: [0, 1440)
          intervals.push({ start: 0, end: 1440 });
        } else if (rType === '遅刻') {
          // 遅刻: [shiftStart, newStart) を不可
          var newStart = rec.startTime;
          if (newStart != null && shift.shiftStartMin != null) {
            intervals.push({ start: shift.shiftStartMin, end: newStart });
          }
        } else if (rType === '早退') {
          // 早退: [newEnd, shiftEnd) を不可
          var newEnd = rec.endTime;
          if (newEnd != null && shift.shiftEndMin != null) {
            intervals.push({ start: newEnd, end: shift.shiftEndMin });
          }
        } else if (rType === '時間指定') {
          // 時間指定: [start, end) を不可
          if (rec.startTime != null && rec.endTime != null) {
            intervals.push({ start: rec.startTime, end: rec.endTime });
          }
        } else if (rType === '午前休') {
          // 午前休: [shiftStart, 12:00) = 12:00まで不可
          var amStart = shift.shiftStartMin != null ? shift.shiftStartMin : 0;
          intervals.push({ start: amStart, end: 720 });
        } else if (rType === '午後休') {
          // 午後休: [12:00, shiftEnd) = 12:00以降不可
          var pmEnd = shift.shiftEndMin != null ? shift.shiftEndMin : 1440;
          intervals.push({ start: 720, end: pmEnd });
        }
      });
    }

    // ※イベントは不可区間に入れない（アンカー方式の後段で処理するため）
    // イベントがあるスタッフでも、その前後に訪問を詰められる可能性があるため
    // 割当段階で候補を潰さず、時刻自動調整フェーズでイベントを避けて配置する

    // 区間がなければ空配列
    if (intervals.length === 0) return [];

    // 区間をマージ（重複・連続区間の統合）
    if (intervals.length <= 1) return intervals;
    intervals.sort(function(a, b) { return a.start - b.start; });
    var merged = [intervals[0]];
    for (var i = 1; i < intervals.length; i++) {
      var last = merged[merged.length - 1];
      var curr = intervals[i];
      if (curr.start <= last.end) {
        // 重複または連続 → マージ
        last.end = Math.max(last.end, curr.end);
      } else {
        merged.push(curr);
      }
    }
    return merged;
  }

  // 2つの区間が重なるかチェック
  function intervalsOverlap_(a, b) {
    return a.start < b.end && b.start < a.end;
  }

  // 固定訪問: 訪問区間が不可区間と1分でも重なれば不可
  function isFixedVisitBlocked_(visitStart, visitEnd, blockedIntervals) {
    if (visitStart == null || visitEnd == null) return false;
    var visitInterval = { start: visitStart, end: visitEnd };
    for (var i = 0; i < blockedIntervals.length; i++) {
      if (intervalsOverlap_(visitInterval, blockedIntervals[i])) {
        return true;  // 衝突あり
      }
    }
    return false;  // 衝突なし
  }

  // 可動訪問: 許容範囲から不可区間を引いた空き区間にsvcMinを置けるか
  function isFlexibleVisitBlocked_(earliestMin, latestMin, svcMin, blockedIntervals) {
    if (earliestMin == null || latestMin == null) return false;
    if (svcMin <= 0) svcMin = 30;  // デフォルト30分

    // 許容範囲 [earliestMin, latestMin] から不可区間を除いた空き区間を計算
    var available = [{ start: earliestMin, end: latestMin }];

    blockedIntervals.forEach(function(blocked) {
      var newAvailable = [];
      available.forEach(function(avail) {
        if (blocked.end <= avail.start || blocked.start >= avail.end) {
          // 重ならない
          newAvailable.push(avail);
        } else {
          // 重なる → 分割
          if (avail.start < blocked.start) {
            newAvailable.push({ start: avail.start, end: blocked.start });
          }
          if (blocked.end < avail.end) {
            newAvailable.push({ start: blocked.end, end: avail.end });
          }
        }
      });
      available = newAvailable;
    });

    // 空き区間のどこかにsvcMinが収まるかチェック
    for (var i = 0; i < available.length; i++) {
      var gap = available[i].end - available[i].start;
      if (gap >= svcMin) {
        return false;  // 収まる → ブロックされていない
      }
    }
    return true;  // どこにも収まらない → ブロック
  }

  // スタッフがこの訪問に対応可能かチェック（スタッフ個別変更を考慮）
  function isStaffAvailableForVisit_(staffId, dateStr, timeType, startMin, endMin, earliestMin, latestMin, svcMin) {
    var blockedIntervals = getStaffBlockedIntervals_(staffId, dateStr);
    if (blockedIntervals.length === 0) return true;  // 制限なし

    // デバッグログ
    console.log('Staff restriction check:', staffId, dateStr, 'blocked:', JSON.stringify(blockedIntervals), 'timeType:', timeType, 'start:', startMin, 'end:', endMin);

    // 終日不可チェック（[0,1440)が含まれていれば完全除外）
    for (var i = 0; i < blockedIntervals.length; i++) {
      if (blockedIntervals[i].start === 0 && blockedIntervals[i].end >= 1440) {
        return false;  // 終日不可
      }
    }

    // 固定訪問または具体的な時間が指定されている場合
    if (timeType === '固定' || (startMin != null && endMin != null)) {
      // 訪問区間が不可区間と重なれば不可
      if (startMin != null && endMin != null) {
        var isBlocked = isFixedVisitBlocked_(startMin, endMin, blockedIntervals);
        console.log('Fixed visit check:', startMin, '-', endMin, 'blocked:', isBlocked);
        return !isBlocked;
      }
    }

    // 可動訪問（午前/午後/終日/時間帯）
    // 許容範囲を決定
    var effEarliest = earliestMin;
    var effLatest = latestMin;

    // timeTypeによるデフォルト許容範囲
    if (timeType === '午前') {
      if (effEarliest == null) effEarliest = 9 * 60;   // 09:00
      if (effLatest == null) effLatest = 12 * 60;      // 12:00
    } else if (timeType === '午後') {
      if (effEarliest == null) effEarliest = 13 * 60;  // 13:00
      if (effLatest == null) effLatest = 17 * 60;      // 17:00
    } else if (timeType === '終日') {
      if (effEarliest == null) effEarliest = 9 * 60;   // 09:00
      if (effLatest == null) effLatest = 18 * 60;      // 18:00
    } else {
      // 時間帯など: start/endから許容範囲を取得
      if (effEarliest == null) effEarliest = startMin;
      if (effLatest == null) effLatest = endMin;
    }

    // 判定不能の場合、スタッフのシフト全体で判定
    if (effEarliest == null || effLatest == null) {
      var shift = getStaffShift_(staffId);
      effEarliest = shift.shiftStartMin || 0;
      effLatest = shift.shiftEndMin || 1440;
    }

    var isBlocked = isFlexibleVisitBlocked_(effEarliest, effLatest, svcMin || 30, blockedIntervals);
    console.log('Flexible visit check:', effEarliest, '-', effLatest, 'svcMin:', svcMin, 'blocked:', isBlocked);
    return !isBlocked;
  }

  var assignCountMap = {};
  function getAssignCount(staffId, dateStr) { return assignCountMap[staffId + '|' + dateStr] || 0; }
  function incAssignCount(staffId, dateStr) { var k = staffId + '|' + dateStr; assignCountMap[k] = (assignCountMap[k] || 0) + 1; }

  var patientWeekCount = {};
  function getPatientWeekCount(pid, staffId) { return patientWeekCount[pid + '|' + staffId] || 0; }
  function incPatientWeekCount(pid, staffId) { var k = pid + '|' + staffId; patientWeekCount[k] = (patientWeekCount[k] || 0) + 1; }

  var weeklyRequests = [];
  wData.forEach(function(row){
    var d = row[wIdx.date];
    if (!(d instanceof Date)) return;
    var changeType = row[wIdx.change];
    if (changeType === 'キャンセル') return;
    weeklyRequests.push({ row: row, date: d, dateStr: Utilities.formatDate(d, tz, 'yyyy/MM/dd'), start: row[wIdx.start], end: row[wIdx.end] });
  });

  weeklyRequests.sort(function(a,b){
    if (a.date.getTime() !== b.date.getTime()) return a.date - b.date;
    return toMinutes(a.start) - toMinutes(b.start);
  });

  var resultRows = [];
  var unassignedList = [];

  // ============================================================
  // Task A-2.5: イベントを resultRows に先挿入（イベント優先の固定アンカー）
  // ============================================================

  // staff_id|dateStr ごとに、イベントの開始/終了を確定させる
  function resolveEventInterval_(ev, staffId) {
    var shift = getStaffShift_(staffId);
    var shiftStart = (shift.shiftStartMin != null ? shift.shiftStartMin : 540);
    var shiftEnd   = (shift.shiftEndMin   != null ? shift.shiftEndMin   : 1080);

    // 1) start/end が両方あるなら最優先
    if (ev.startMin != null && ev.endMin != null && ev.endMin > ev.startMin) {
      return { start: ev.startMin, end: ev.endMin };
    }

    // 2) timeMode + duration で決める
    var dur = Number(ev.durationMin || 60) || 60;
    var mode = String(ev.timeMode || '').trim();

    if (mode === '午前') {
      var s = shiftStart;
      var e = Math.min(s + dur, 720);
      return { start: s, end: e };
    }
    if (mode === '午後') {
      var s2 = 720;
      var e2 = Math.min(s2 + dur, shiftEnd);
      return { start: s2, end: e2 };
    }
    if (mode === '終日') {
      return { start: shiftStart, end: shiftEnd };
    }

    // 3) それでも決まらない → 仮置き（目印用：13:00開始）
    var s3 = Math.max(shiftStart, 13 * 60);
    var e3 = Math.min(s3 + dur, shiftEnd);
    return { start: s3, end: e3 };
  }

  // イベント行をresultRowsへ
  Object.keys(eventMap).forEach(function(key) {
    var parts = key.split('|');
    var staffId = parts[0];
    var dateStr = parts[1];

    var events = eventMap[key] || [];
    if (!staffId || !dateStr || events.length === 0) return;

    // dateStr -> Date化
    var dateObj = parseDate_(dateStr);
    if (!dateObj) return;

    events.forEach(function(ev, idxEv) {
      var itv = resolveEventInterval_(ev, staffId);
      if (!itv || itv.start == null || itv.end == null || itv.end <= itv.start) return;

      var evid = ev.eventId ? String(ev.eventId) : ('EVT' + ('000' + (idxEv + 1)).slice(-3));
      var visitId = 'EV_' + evid;

      var evLabel = '[EV] ' + (ev.eventType || '') + (ev.title ? ':' + ev.title : '');
      var note = evLabel + (ev.patientLinked && ev.patientId ? ('（' + ev.patientId + '）') : '');

      // スタッフ名を取得
      var staffName = '';
      for (var si = 0; si < staffList.length; si++) {
        if (staffList[si].id === staffId) { staffName = staffList[si].name; break; }
      }

      // resultRowsのフォーマット:
      // [visit_id, 日付, 曜日, staff_id, スタッフ名, patient_id, 患者名, エリア, 開始時刻, 終了時刻,
      //  サービス時間, 時間タイプ, 希望最早時刻, 希望最遅時刻, 備考]
      resultRows.push([
        visitId,
        dateObj,
        '',               // 曜日
        staffId,
        staffName,
        '',               // patient_id
        '',               // 患者名
        'EV',             // エリア（目印）
        minutesToSerial_(itv.start),
        minutesToSerial_(itv.end),
        (itv.end - itv.start), // サービス時間(分)
        '固定',
        minutesToSerial_(itv.start), // 希望最早
        minutesToSerial_(itv.end),   // 希望最遅
        note
      ]);
    });
  });

  // ============================================================
  // 割当時の重なりチェック用：staffDateMapを先に初期化（イベント含む）
  // ============================================================
  var staffDateMap = {};
  resultRows.forEach(function(r, i) {
    var d = r[1], staffId = r[3];
    if (!staffId || !(d instanceof Date)) return;
    var key = staffId + '|' + Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    if (!staffDateMap[key]) staffDateMap[key] = [];
    staffDateMap[key].push(i);
  });

  // 固定訪問が既存予定と重なるかチェックするヘルパー
  function isOverlappedWithExisting_(staffId, dateObj, startMin, endMin) {
    if (!staffId || !(dateObj instanceof Date) || startMin == null || endMin == null) return false;

    var dateStr = Utilities.formatDate(dateObj, tz, 'yyyy/MM/dd');
    var key = staffId + '|' + dateStr;
    var idxList = staffDateMap[key] || [];

    for (var i = 0; i < idxList.length; i++) {
      var r = resultRows[idxList[i]];
      if (!r) continue;
      // 未割当は無視
      if (!r[3] || r[4] === '未割当') continue;

      var s = toMinutes(r[8]);
      var e = toMinutes(r[9]);
      if (s == null || e == null) continue;

      if (s < endMin && startMin < e) return true; // overlap
    }
    return false;
  }

  weeklyRequests.forEach(function(item, idx){
    var row = item.row;
    var dateObj = item.date;
    var dateStr = item.dateStr;
    var youbiRaw = row[wIdx.youbi];
    var youbi = normalizeYoubi(youbiRaw);
    var pid = row[wIdx.pid];
    var pname = row[wIdx.pname];
    var area = row[wIdx.area];
    var start = row[wIdx.start];
    var end = row[wIdx.end];

    var svcRaw = row[wIdx.svcMin];
    var svcMin = Number(svcRaw);
    if (!svcMin && patientMap[pid] && patientMap[pid].svcMin) svcMin = Number(patientMap[pid].svcMin) || 0;
    if (!svcMin && typeof svcRaw === 'string') { var m = svcRaw.match(/(\d+)/); if (m) svcMin = Number(m[1]); }
    if (!svcMin) svcMin = 0;

    var sexLimit = row[wIdx.sexLimit];
    var contPrefRaw = row[wIdx.contPref];
    var contPref = normalizeContPref(contPrefRaw);
    var prevSid = row[wIdx.prevSid];
    var prevSname = row[wIdx.prevSname];
    var note = row[wIdx.note];
    var timeType = row[wIdx.timeType];
    var earliest = row[wIdx.earliest];
    var latest = row[wIdx.latest];

    var specifiedIdsArr = wIdx.specifiedIds >= 0 ? parseIdList(row[wIdx.specifiedIds]) : [];
    var specifiedType = wIdx.specifiedType >= 0 ? String(row[wIdx.specifiedType] || '').trim() : '';
    var ngIdsArr = wIdx.ngStaffIds >= 0 ? parseIdList(row[wIdx.ngStaffIds]) : [];

    var startMin = toMinutes(start);
    var endMin = toMinutes(end);

    var pInfo = patientMap[pid] || {};
    var plat = pInfo.lat;
    var plng = pInfo.lng;

    var avoidPrev = (contPref === 'ローテーション優先' && prevSid);
    var earliestMin = earliest ? toMinutes(earliest) : null;
    var latestMin = latest ? toMinutes(latest) : null;

    function canStaffServe(st, preferAreaFlagObj) {
      if (sexLimit === '女性のみ' && st.gender !== '女性') return false;
      if (sexLimit === '男性のみ' && st.gender !== '男性') return false;
      if (youbi && st.workDays.length > 0 && st.workDays.indexOf(youbi) === -1) return false;
      if (st.shiftStartMin != null && st.shiftEndMin != null) {
        if (timeType === '固定') {
          if (startMin != null && startMin < st.shiftStartMin) return false;
          if (endMin != null && endMin > st.shiftEndMin) return false;
          // ★追加：同スタッフの既存予定と重なるなら不可
          if (startMin != null && endMin != null) {
            if (isOverlappedWithExisting_(st.id, dateObj, startMin, endMin)) return false;
          }
        } else {
          var reqStart = earliestMin != null ? earliestMin : startMin;
          var reqEnd = latestMin != null ? latestMin : endMin;
          if (reqStart == null) reqStart = st.shiftStartMin;
          if (reqEnd == null) reqEnd = st.shiftEndMin;
          var latestStart = Math.max(reqStart, st.shiftStartMin);
          var earliestEnd = Math.min(reqEnd, st.shiftEndMin);
          if (latestStart >= earliestEnd) return false;
        }
      }
      var count = getAssignCount(st.id, dateStr);
      if (count >= st.maxPerDay) return false;

      // スタッフ個別変更リクエストによる制限チェック
      if (!isStaffAvailableForVisit_(st.id, dateStr, timeType, startMin, endMin, earliestMin, latestMin, svcMin)) {
        return false;
      }

      preferAreaFlagObj.flag = false;
      return true;
    }

    var needStaff = wIdx.needStaff >= 0 ? toHalfWidthNumber_(row[wIdx.needStaff], 1) : 1;
    if (needStaff < 1) needStaff = 1;
    if (needStaff > 2) needStaff = 2;

    var usedStaffIds = {};

    for (var slot = 1; slot <= needStaff; slot++) {
      var chosenStaff = null;

      if (specifiedType === '必須' && specifiedIdsArr.length > 0) {
        for (var si = 0; si < specifiedIdsArr.length; si++) {
          var specId = specifiedIdsArr[si];
          if (usedStaffIds[specId]) continue;
          var stSpec = staffList.find(function(s){ return s.id === specId; });
          if (stSpec) {
            var objSpec = { flag: false };
            if (canStaffServe(stSpec, objSpec)) { chosenStaff = stSpec; break; }
          }
        }
        if (!chosenStaff) note = (note || '') + ' / 指定必須スタッフ割当不可';
      } else if (contPref === '同じ人希望' && prevSid && !usedStaffIds[prevSid]) {
        if (ngIdsArr.indexOf(prevSid) < 0) {
          var stPrev = staffList.find(function(s){ return s.id === prevSid; });
          if (stPrev) {
            var objPrev = { flag: false };
            if (canStaffServe(stPrev, objPrev)) chosenStaff = stPrev;
          }
        }
      }

      if (!chosenStaff) {
        var candidates = [];
        staffList.forEach(function(st){
          if (usedStaffIds[st.id]) return;
          if (sexLimit === '女性のみ' && st.gender !== '女性') return;
          if (sexLimit === '男性のみ' && st.gender !== '男性') return;
          if (youbi && st.workDays.length > 0 && st.workDays.indexOf(youbi) === -1) return;
          if (avoidPrev && prevSid && st.id === prevSid) return;
          if (st.shiftStartMin != null && st.shiftEndMin != null) {
            if (timeType === '固定') {
              if (startMin != null && startMin < st.shiftStartMin) return;
              if (endMin != null && endMin > st.shiftEndMin) return;
              // ★追加：同スタッフの既存予定と重なるなら不可
              if (startMin != null && endMin != null) {
                if (isOverlappedWithExisting_(st.id, dateObj, startMin, endMin)) return;
              }
            } else {
              var reqStart = earliestMin != null ? earliestMin : startMin;
              var reqEnd = latestMin != null ? latestMin : endMin;
              if (reqStart == null) reqStart = st.shiftStartMin;
              if (reqEnd == null) reqEnd = st.shiftEndMin;
              var latestStart = Math.max(reqStart, st.shiftStartMin);
              var earliestEnd = Math.min(reqEnd, st.shiftEndMin);
              if (latestStart >= earliestEnd) return;
            }
          }
          var dayCount = getAssignCount(st.id, dateStr);
          if (dayCount >= st.maxPerDay) return;

          // スタッフ個別変更リクエストによる制限チェック
          if (!isStaffAvailableForVisit_(st.id, dateStr, timeType, startMin, endMin, earliestMin, latestMin, svcMin)) return;

          var distKm = calcDistanceKm(plat, plng, st.lat, st.lng);
          candidates.push({ staff: st, dayCount: dayCount, patientCount: getPatientWeekCount(pid, st.id),
                           distKm: distKm, distScore: distToScore(distKm), samePatientToday: false });
        });

        candidates = applyStaffPreference(candidates, specifiedIdsArr, specifiedType, ngIdsArr);

        if (candidates.length > 0) {
          candidates.forEach(function(c){
            var samePatientToday = resultRows.some(function(rr){
              return rr[3] === c.staff.id && rr[5] === pid && Utilities.formatDate(rr[1], tz, 'yyyy/MM/dd') === dateStr;
            });
            c.samePatientToday = samePatientToday;
          });

          candidates.sort(function(a, b){
            if (a._pref !== undefined && b._pref !== undefined && a._pref !== b._pref) return b._pref - a._pref;
            if (a.samePatientToday !== b.samePatientToday) return a.samePatientToday ? -1 : 1;
            if (contPref === 'ローテーション優先' && a.patientCount !== b.patientCount) return a.patientCount - b.patientCount;
            if (a.distScore !== b.distScore) return a.distScore - b.distScore;
            return a.dayCount - b.dayCount;
          });
          chosenStaff = candidates[0].staff;
        }
      }

      if (!chosenStaff) {
        var fallback = [];
        staffList.forEach(function(st){
          if (usedStaffIds[st.id]) return;
          if (ngIdsArr.indexOf(st.id) >= 0) return;
          if (sexLimit === '女性のみ' && st.gender !== '女性') return;
          if (sexLimit === '男性のみ' && st.gender !== '男性') return;
          if (youbi && st.workDays.length > 0 && st.workDays.indexOf(youbi) === -1) return;
          if (startMin != null && st.shiftStartMin != null && startMin < st.shiftStartMin) return;
          if (endMin != null && st.shiftEndMin != null && endMin > st.shiftEndMin) return;

          // スタッフ個別変更リクエストによる制限チェック（fallback時も適用）
          if (!isStaffAvailableForVisit_(st.id, dateStr, timeType, startMin, endMin, earliestMin, latestMin, svcMin)) return;

          fallback.push({ staff: st, dayCount: getAssignCount(st.id, dateStr) });
        });
        if (fallback.length > 0) {
          fallback.sort(function(a,b){ return a.dayCount - b.dayCount; });
          chosenStaff = fallback[0].staff;
          note = (note || '') + ' / 自動割当: 上限超過の可能性あり';
        }
      }

      var staffId = '', staffName = '';
      if (chosenStaff) {
        staffId = chosenStaff.id;
        staffName = chosenStaff.name;
        usedStaffIds[staffId] = true;
        incAssignCount(staffId, dateStr);
        incPatientWeekCount(pid, staffId);
      } else {
        staffName = '未割当';
        unassignedList.push({ date: dateObj, youbi: youbiRaw, pid: pid, pname: pname, needStaff: needStaff, slot: slot, reason: note || '条件を満たすスタッフなし' });
      }

      var visitId = 'V' + Utilities.formatString('%03d', idx + 1);
      if (needStaff > 1) visitId = visitId + '-' + slot;

      var note2 = note || '';
      if (needStaff > 1) note2 = (note2 ? note2 + ' / ' : '') + '同時訪問(' + slot + '/' + needStaff + ')';

      resultRows.push([visitId, dateObj, youbiRaw, staffId, staffName, pid, pname, area, start, end, svcMin, timeType, earliest, latest, note2]);

      // ★追加：この行を staffDateMap に登録（後続の重なり判定に使う）
      if (staffId && dateObj instanceof Date) {
        var k = staffId + '|' + dateStr;
        if (!staffDateMap[k]) staffDateMap[k] = [];
        staffDateMap[k].push(resultRows.length - 1);
      }
    }
  });

  // ルート最適化（簡略版）
  var staffLocMap = {};
  staffList.forEach(function(st){ staffLocMap[st.id] = { lat: Number(st.lat) || null, lng: Number(st.lng) || null }; });

  var dayGroupMap = {};
  resultRows.forEach(function(r, i){
    var d = r[1], sId = r[3];
    if (!sId || !(d instanceof Date)) return;
    var key = sId + '|' + Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    if (!dayGroupMap[key]) dayGroupMap[key] = [];
    dayGroupMap[key].push(i);
  });

  Object.keys(dayGroupMap).forEach(function(key){
    var idxList = dayGroupMap[key];
    idxList.sort(function(aIdx, bIdx){ return toMinutes(resultRows[aIdx][8]) - toMinutes(resultRows[bIdx][8]); });
  });

  // 移動距離計算
  var prevVisitIdArr = new Array(resultRows.length).fill('');
  var moveKmArr = new Array(resultRows.length).fill('');
  var moveMinArr = new Array(resultRows.length).fill('');

  // staffDateMapを再構築（時刻調整・Level1再挿入用：イベント含む）
  staffDateMap = {};
  resultRows.forEach(function(r, i){
    var d = r[1], staffId = r[3];
    // イベント行も含める（時刻調整のアンカー、Level1再挿入の隙間判定に必要）
    if (!staffId || !(d instanceof Date)) return;
    var key = staffId + '|' + Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    if (!staffDateMap[key]) staffDateMap[key] = [];
    staffDateMap[key].push(i);
  });

  // ★移動距離計算専用：患者行のみ（イベントはpid空なので除外）
  var staffDateMapForMove = {};
  resultRows.forEach(function(r, i){
    var d = r[1], staffId = r[3], pid = r[5];
    if (!staffId || !(d instanceof Date) || !pid) return;
    var key = staffId + '|' + Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    if (!staffDateMapForMove[key]) staffDateMapForMove[key] = [];
    staffDateMapForMove[key].push(i);
  });

  Object.keys(staffDateMapForMove).forEach(function(key){
    var indexList = staffDateMapForMove[key];
    indexList.sort(function(aIdx, bIdx){ return toMinutes(resultRows[aIdx][8]) - toMinutes(resultRows[bIdx][8]); });
    for (var j = 0; j < indexList.length; j++) {
      var currIndex = indexList[j];
      if (j === 0) continue;
      var prevIndex = indexList[j - 1];
      var prevPid = resultRows[prevIndex][5];
      var currPid = resultRows[currIndex][5];
      var prevP = patientMap[prevPid] || {};
      var currP = patientMap[currPid] || {};
      if (prevP.lat != null && prevP.lng != null && currP.lat != null && currP.lng != null) {
        var distKm = calcDistanceKm(prevP.lat, prevP.lng, currP.lat, currP.lng);
        prevVisitIdArr[currIndex] = resultRows[prevIndex][0];
        moveKmArr[currIndex] = distKm;
        moveMinArr[currIndex] = Math.round(distKm / 20 * 60);
      }
    }
  });

  // ============================================================
  // 時刻自動調整（イベント/固定訪問をアンカーとして可動訪問を隙間に詰める）
  // ============================================================

  // "その行がイベントか" 判定（visit_idがEV_ か 備考に[EV]）
  function isEventRow_(row) {
    var vid = String(row[0] || '');
    var note = String(row[14] || '');
    return vid.indexOf('EV_') === 0 || note.indexOf('[EV]') >= 0;
  }

  // 分ベースのstart/end取得
  function rowStartMin_(row) { return toMinutes(row[8]); }
  function rowEndMin_(row)   { return toMinutes(row[9]); }

  function setRowTimeByMinutes_(row, startMin, endMin) {
    row[8] = minutesToSerial_(startMin);
    row[9] = minutesToSerial_(endMin);
  }

  // 2区間が重なるか
  function overlap_(s1,e1,s2,e2) {
    return s1 < e2 && s2 < e1;
  }

  // 固定区間を避けて次の開始時刻を見つける
  function findNextNonOverlappingStart_(candidateStart, durationMin, fixedIntervals, latestMin) {
    var s = candidateStart;
    var e = s + durationMin;

    // ガード: latestMin制約チェック
    if (latestMin != null && s > latestMin) return null;

    // 固定区間と重なる限り、固定区間の end までジャンプ
    var safety = 0;
    while (safety < 200) {
      safety++;

      var hit = null;
      for (var i = 0; i < fixedIntervals.length; i++) {
        var f = fixedIntervals[i];
        if (overlap_(s, e, f.start, f.end)) { hit = f; break; }
      }
      if (!hit) return s; // 重ならない → この開始時刻でOK

      s = hit.end; // 固定区間の終了後に移動
      e = s + durationMin;

      if (latestMin != null && e > latestMin) return null; // 時間切れ
    }
    return null; // ループ防止（見つからない）
  }

  Object.keys(staffDateMap).forEach(function(key){
    var idxList = staffDateMap[key];
    if (!idxList || idxList.length === 0) return;

    // 時刻未確定が混ざるので、まずアンカー/可動で分ける
    var anchors = []; // {idx, s, e, kind}
    var flexes  = []; // idx

    idxList.forEach(function(rIdx){
      var row = resultRows[rIdx];
      var timeType = row[11];
      var s = rowStartMin_(row);
      var e = rowEndMin_(row);

      var isEv = isEventRow_(row);

      // アンカー条件：
      // - イベント（常に固定）
      // - timeType固定 かつ start/end がある（固定訪問）
      if (isEv) {
        if (s != null && e != null && e > s) anchors.push({ idx: rIdx, s: s, e: e, kind: 'EV' });
        else {
          // イベントなのに時間が取れない → とりあえず何もしない（要データ）
          row[14] = (row[14] || '') + ' / EV時間未確定';
          flexes.push(rIdx);
        }
        return;
      }

      if (timeType === '固定' && s != null && e != null && e > s) {
        anchors.push({ idx: rIdx, s: s, e: e, kind: 'FIX' });
        return;
      }

      // それ以外は可動として扱う
      flexes.push(rIdx);
    });

    anchors.sort(function(a,b){ return a.s - b.s; });

    // ============================================================
    // アンカー同士（EV/FIX）の衝突も解消：後勝ちを未割当に落とす
    // ルール：EVは最優先。EVと衝突したFIXは落とす。FIX同士なら後の方を落とす。
    // ============================================================
    (function resolveAnchorConflicts_(){
      anchors.sort(function(a,b){ return a.s - b.s; });

      var cleaned = [];
      for (var i = 0; i < anchors.length; i++) {
        var cur = anchors[i];
        if (cleaned.length === 0) { cleaned.push(cur); continue; }

        var prev = cleaned[cleaned.length - 1];

        // 重なり判定
        if (overlap_(prev.s, prev.e, cur.s, cur.e)) {
          // EV優先：FIXを落とす
          if (prev.kind === 'EV' && cur.kind === 'FIX') {
            // cur を落とす
            var rowC = resultRows[cur.idx];
            rowC[3] = '';
            rowC[4] = '未割当';
            rowC[14] = (rowC[14] || '') + ' / アンカー衝突(EV優先で未割当)';
            unassignedList.push({ date: rowC[1], youbi: rowC[2], pid: rowC[5], pname: rowC[6], needStaff: 1, slot: 1, reason: 'アンカー衝突(EV優先)' });
            continue;
          }
          if (prev.kind === 'FIX' && cur.kind === 'EV') {
            // prev を落として EV を採用（cleaned の最後を置き換え）
            var rowP = resultRows[prev.idx];
            rowP[3] = '';
            rowP[4] = '未割当';
            rowP[14] = (rowP[14] || '') + ' / アンカー衝突(EV優先で未割当)';
            unassignedList.push({ date: rowP[1], youbi: rowP[2], pid: rowP[5], pname: rowP[6], needStaff: 1, slot: 1, reason: 'アンカー衝突(EV優先)' });

            cleaned[cleaned.length - 1] = cur;
            continue;
          }

          // FIX同士：後勝ち（cur）を落とす
          if (prev.kind === 'FIX' && cur.kind === 'FIX') {
            var rowF = resultRows[cur.idx];
            rowF[3] = '';
            rowF[4] = '未割当';
            rowF[14] = (rowF[14] || '') + ' / 固定同士衝突で未割当';
            unassignedList.push({ date: rowF[1], youbi: rowF[2], pid: rowF[5], pname: rowF[6], needStaff: 1, slot: 1, reason: '固定同士が衝突' });
            continue;
          }

          // EV同士が衝突してる場合はデータ側問題（どちらも残す、要確認を記録）
          if (prev.kind === 'EV' && cur.kind === 'EV') {
            var rowE = resultRows[cur.idx];
            rowE[14] = (rowE[14] || '') + ' / EV同士衝突(要確認)';
            // 両方残す（どちらも固定なので）
            cleaned.push(cur);
            continue;
          }
        }

        cleaned.push(cur);
      }

      anchors = cleaned;
    })();

    // スタッフのシフト範囲を取得
    var staffId = key.split('|')[0];
    var dateStr = key.split('|')[1];
    var shift = getStaffShift_(staffId);
    var dayStart = (shift.shiftStartMin != null ? shift.shiftStartMin : 540);
    var dayEnd   = (shift.shiftEndMin   != null ? shift.shiftEndMin   : 1080);

    // 可動訪問を「隙間」に詰める
    flexes.sort(function(aIdx,bIdx){
      var aS = rowStartMin_(resultRows[aIdx]);
      var bS = rowStartMin_(resultRows[bIdx]);
      if (aS == null && bS == null) return 0;
      if (aS == null) return 1;
      if (bS == null) return -1;
      return aS - bS;
    });

    // gapを作る（アンカー前後にバッファを確保）
    var gaps = [];
    var cursor = dayStart;
    anchors.forEach(function(a){
      // アンカー開始前にバッファを空ける
      var gapEnd = a.s - EXTRA_BUFFER_MIN;
      if (cursor < gapEnd) gaps.push({ s: cursor, e: gapEnd });
      // アンカー終了後にバッファを空けて次へ
      cursor = Math.max(cursor, a.e + EXTRA_BUFFER_MIN);
    });
    if (cursor < dayEnd) gaps.push({ s: cursor, e: dayEnd });

    // gap内へ順に配置
    var cursorByGap = gaps.map(function(g){ return g.s; });

    for (var fi = 0; fi < flexes.length; fi++) {
      var rIdx = flexes[fi];
      var row = resultRows[rIdx];

      // 既に未割当になってるものはスキップ
      if (String(row[4] || '') === '未割当') continue;

      var svcMin = Number(row[10]) || 0;
      if (!svcMin) svcMin = 30;

      // 希望時間帯
      var earliestMin = toMinutes(row[12]);
      var latestMin   = toMinutes(row[13]);

      var placed = false;

      for (var gi = 0; gi < gaps.length; gi++) {
        var g = gaps[gi];
        var startCand = cursorByGap[gi];

        // 希望がある場合はその範囲に寄せる
        if (earliestMin != null) startCand = Math.max(startCand, earliestMin);
        var endCand = startCand + svcMin;

        // 最新制約（latestMin は「終了上限」として扱う）
        if (latestMin != null && endCand > latestMin) {
          // このgapでは無理かも → 次のgapへ
          continue;
        }

        // gap内に収まるか
        if (startCand >= g.s && endCand <= g.e) {
          setRowTimeByMinutes_(row, startCand, endCand);
          cursorByGap[gi] = endCand + EXTRA_BUFFER_MIN; // 次はバッファ込みで進める
          placed = true;
          break;
        }
      }

      if (!placed) {
        // 置けない → 未割当化
        row[3] = '';
        row[4] = '未割当';
        row[14] = (row[14] || '') + ' / EV優先: 隙間に入らず未割当';
        unassignedList.push({
          date: row[1],
          youbi: row[2],
          pid: row[5],
          pname: row[6],
          needStaff: 1,
          slot: 1,
          reason: 'イベント優先で隙間に入らない'
        });
      }
    }
  });

  // ============================================================
  // Level 1: 未割当訪問の再挿入（別スタッフへの振り分け試行）
  // ============================================================

  // 未割当になった訪問を抽出（staff_idが空 or スタッフ名が"未割当"の行）
  var unassignedRows = [];
  for (var ui = 0; ui < resultRows.length; ui++) {
    var uRow = resultRows[ui];
    if (!uRow[3] || uRow[4] === '未割当') {
      // イベント行は除外
      if (String(uRow[0] || '').indexOf('EV_') === 0) continue;
      unassignedRows.push({ idx: ui, row: uRow });
    }
  }

  if (unassignedRows.length > 0) {
    // 候補スタッフを選定する関数
    function getCandidateStaffsForVisit_(visitRow, dateStr, youbi) {
      var candidates = [];
      var pid = visitRow[5];
      var pInfo = patientMap[pid] || {};
      var pLat = pInfo.lat;
      var pLng = pInfo.lng;

      for (var si = 0; si < staffList.length; si++) {
        var staff = staffList[si];

        // 1) 曜日チェック
        if (staff.workDays && staff.workDays.length > 0) {
          if (staff.workDays.indexOf(youbi) < 0) continue;
        }

        // 2) スタッフ変更（終日休み等）チェック
        var scKey = staff.id + '|' + dateStr;
        var scRecs = staffChangeMap[scKey] || [];
        var isDayOff = false;
        for (var sci = 0; sci < scRecs.length; sci++) {
          var rt = scRecs[sci].restrictionType;
          if (rt === '休み' || rt === '終日不可' || rt === '終日') {
            isDayOff = true;
            break;
          }
        }
        if (isDayOff) continue;

        // 3) 距離でスコア（近い順）
        var distScore = 9999;
        if (pLat != null && pLng != null && staff.lat != null && staff.lng != null) {
          distScore = calcDistanceKm(pLat, pLng, staff.lat, staff.lng);
        }

        candidates.push({ staff: staff, distScore: distScore });
      }

      // 距離近い順にソート
      candidates.sort(function(a, b) { return a.distScore - b.distScore; });

      // 上位N人に絞る
      return candidates.slice(0, 10);
    }

    // スタッフ+日の隙間に訪問が入るかチェック
    function canInsertVisitToStaffDay_(staffId, dateStr, visitRow, svcMin) {
      var dayKey = staffId + '|' + dateStr;
      var dayIdxList = staffDateMap[dayKey] || [];
      if (dayIdxList.length === 0) {
        // そのスタッフのその日に何もなければ、シフト内なら入る
        var shift = getStaffShift_(staffId);
        var shiftS = shift.shiftStartMin != null ? shift.shiftStartMin : 540;
        var shiftE = shift.shiftEndMin != null ? shift.shiftEndMin : 1080;
        return (shiftE - shiftS >= svcMin) ? { ok: true, startMin: shiftS } : { ok: false };
      }

      // 既存の予定を集める
      var anchors = [];
      for (var di = 0; di < dayIdxList.length; di++) {
        var rIdx = dayIdxList[di];
        var r = resultRows[rIdx];
        var s = toMinutes(r[8]);
        var e = toMinutes(r[9]);
        if (s != null && e != null && e > s) {
          anchors.push({ start: s, end: e });
        }
      }
      anchors.sort(function(a, b) { return a.start - b.start; });

      // 隙間を探す
      var shift2 = getStaffShift_(staffId);
      var dayStart = shift2.shiftStartMin != null ? shift2.shiftStartMin : 540;
      var dayEnd = shift2.shiftEndMin != null ? shift2.shiftEndMin : 1080;

      var gaps = [];
      var cursor = dayStart;
      for (var ai = 0; ai < anchors.length; ai++) {
        if (cursor + EXTRA_BUFFER_MIN < anchors[ai].start) {
          gaps.push({ s: cursor, e: anchors[ai].start - EXTRA_BUFFER_MIN });
        }
        cursor = Math.max(cursor, anchors[ai].end + EXTRA_BUFFER_MIN);
      }
      if (cursor < dayEnd) {
        gaps.push({ s: cursor, e: dayEnd });
      }

      // svcMin が入る隙間を探す
      for (var gi = 0; gi < gaps.length; gi++) {
        var gapLen = gaps[gi].e - gaps[gi].s;
        if (gapLen >= svcMin) {
          return { ok: true, startMin: gaps[gi].s };
        }
      }

      return { ok: false };
    }

    // 再挿入ループ
    var reinsertedCount = 0;
    for (var uri = 0; uri < unassignedRows.length; uri++) {
      var uItem = unassignedRows[uri];
      var uRow = uItem.row;
      var uIdx = uItem.idx;

      var uDateObj = uRow[1];
      if (!(uDateObj instanceof Date)) continue;
      var uDateStr = Utilities.formatDate(uDateObj, tz, 'yyyy/MM/dd');
      var uYoubi = normalizeYoubi(uRow[2]);
      var uSvcMin = Number(uRow[10]) || 30;

      var candidates = getCandidateStaffsForVisit_(uRow, uDateStr, uYoubi);

      for (var ci = 0; ci < candidates.length; ci++) {
        var cStaff = candidates[ci].staff;
        var cStaffId = cStaff.id;

        // 元のスタッフは除外（すでに試したはず）
        if (cStaffId === uRow[3]) continue;

        var insertResult = canInsertVisitToStaffDay_(cStaffId, uDateStr, uRow, uSvcMin);
        if (insertResult.ok) {
          // 挿入成功！ rowを更新
          uRow[3] = cStaffId;
          uRow[4] = cStaff.name;
          uRow[8] = minutesToSerial_(insertResult.startMin);
          uRow[9] = minutesToSerial_(insertResult.startMin + uSvcMin);
          uRow[14] = (uRow[14] || '') + ' / 再挿入(' + cStaff.name + ')';

          // staffDateMap に追加
          var newKey = cStaffId + '|' + uDateStr;
          if (!staffDateMap[newKey]) staffDateMap[newKey] = [];
          staffDateMap[newKey].push(uIdx);

          reinsertedCount++;
          break;
        }
      }
    }

    // unassignedList から再挿入成功分を除去
    if (reinsertedCount > 0) {
      var newUnassignedList = [];
      for (var nui = 0; nui < unassignedList.length; nui++) {
        var uItem2 = unassignedList[nui];
        // resultRows で同じ patient_id + date が未割当のままかチェック
        var stillUnassigned = false;
        for (var rri = 0; rri < resultRows.length; rri++) {
          var rr = resultRows[rri];
          if (rr[5] === uItem2.pid && rr[1] instanceof Date) {
            var rrDateStr = Utilities.formatDate(rr[1], tz, 'yyyy/MM/dd');
            var uItem2DateStr = (uItem2.date instanceof Date) ?
              Utilities.formatDate(uItem2.date, tz, 'yyyy/MM/dd') : String(uItem2.date);
            if (rrDateStr === uItem2DateStr && (!rr[3] || rr[4] === '未割当')) {
              stillUnassigned = true;
              break;
            }
          }
        }
        if (stillUnassigned) {
          newUnassignedList.push(uItem2);
        }
      }
      unassignedList = newUnassignedList;
    }
  }

  // ============================================================
  // Level3: 距離最適化（安全版）
  // - EV/固定は動かさない
  // - 希望最早/希望最遅が空の可動訪問のみ並べ替え
  // - 並べ替え後、各「隙間」内で詰め直す（EXTRA_BUFFER_MIN 適用）
  // ============================================================
  function applyLevel3RouteOptimizeSafe_() {

    function getPatientLatLng_(pid) {
      var p = patientMap[pid];
      if (!p) return null;
      if (p.lat == null || p.lng == null) return null;
      return { lat: Number(p.lat), lng: Number(p.lng) };
    }

    function getRowLatLng_(row) {
      var pid = row[5];
      if (!pid) return null;
      return getPatientLatLng_(pid);
    }

    function distKmByLL_(a, b) {
      if (!a || !b) return null;
      return calcDistanceKm(a.lat, a.lng, b.lat, b.lng);
    }

    // 2-optで改善（近い順ベースの経路をちょい改善）
    function twoOpt_(nodes) {
      // nodes: [{idx, ll}]
      if (nodes.length <= 3) return nodes;

      function routeLen(arr) {
        var sum = 0;
        for (var i = 0; i < arr.length - 1; i++) {
          var d = distKmByLL_(arr[i].ll, arr[i+1].ll);
          if (d != null) sum += d;
        }
        return sum;
      }

      var best = nodes.slice();
      var bestLen = routeLen(best);

      var improved = true;
      var guard = 0;
      while (improved && guard < 50) {
        guard++;
        improved = false;

        for (var i = 1; i < best.length - 2; i++) {
          for (var k = i + 1; k < best.length - 1; k++) {
            var cand = best.slice(0, i)
              .concat(best.slice(i, k + 1).reverse())
              .concat(best.slice(k + 1));
            var candLen = routeLen(cand);
            if (candLen + 1e-9 < bestLen) {
              best = cand;
              bestLen = candLen;
              improved = true;
            }
          }
        }
      }
      return best;
    }

    // 近傍法（スタート地点から近い順に作る）
    function nearestNeighbor_(nodes, startLL) {
      if (nodes.length <= 1) return nodes.slice();

      var rest = nodes.slice();
      var route = [];

      var currLL = startLL || rest[0].ll;
      while (rest.length > 0) {
        var bestIdx = 0;
        var bestDist = Infinity;
        for (var i = 0; i < rest.length; i++) {
          var d = distKmByLL_(currLL, rest[i].ll);
          // 位置が取れない場合は後回しになりやすいように大きめ
          var dd = (d == null ? 99999 : d);
          if (dd < bestDist) {
            bestDist = dd;
            bestIdx = i;
          }
        }
        var pick = rest.splice(bestIdx, 1)[0];
        route.push(pick);
        currLL = pick.ll || currLL;
      }
      return route;
    }

    // staffDateMap を使って「スタッフ×日」ごとに処理
    Object.keys(staffDateMap).forEach(function(key){
      var idxList = staffDateMap[key];
      if (!idxList || idxList.length === 0) return;

      var staffId = key.split('|')[0];

      // スタッフ基点（lat/lng）: スタッフマスタの緯度経度
      var staffLL = null;
      for (var si = 0; si < staffList.length; si++) {
        if (staffList[si].id === staffId) {
          var slat = Number(staffList[si].lat);
          var slng = Number(staffList[si].lng);
          if (!isNaN(slat) && !isNaN(slng)) staffLL = { lat: slat, lng: slng };
          break;
        }
      }

      // 予定を「アンカー」と「可動」に分ける
      var anchors = []; // {idx, s, e}
      var flex = [];    // {idx, row, s, e, svcMin, ll}

      idxList.forEach(function(rIdx){
        var row = resultRows[rIdx];
        if (!row) return;

        // 未割当はスキップ
        if (!row[3] || String(row[4] || '') === '未割当') return;

        var isEv = isEventRow_(row);
        var timeType = row[11];
        var s = toMinutes(row[8]);
        var e = toMinutes(row[9]);

        var earliestMin = toMinutes(row[12]);
        var latestMin   = toMinutes(row[13]);

        // アンカー：EV or 固定（時間確定）
        if (isEv || (timeType === '固定' && s != null && e != null && e > s)) {
          if (s != null && e != null && e > s) anchors.push({ idx: rIdx, s: s, e: e });
          return;
        }

        // "安全版"なので、希望最早/最遅があるものは並べ替え対象外（そのまま）
        if (earliestMin != null || latestMin != null) {
          // 触らないが「可動」なので隙間詰め対象にはなり得る
          // ここでは順序最適化から除外するため anchors 側に入れて固定扱いにする
          if (s != null && e != null && e > s) anchors.push({ idx: rIdx, s: s, e: e });
          return;
        }

        var svcMin = Number(row[10]) || 30;
        var ll = getRowLatLng_(row);
        flex.push({ idx: rIdx, row: row, s: s, e: e, svcMin: svcMin, ll: ll });
      });

      if (flex.length <= 2) return; // 少なければ意味が薄い

      // アンカー時刻順
      anchors.sort(function(a,b){ return a.s - b.s; });

      // スタッフのシフト範囲
      var shift = getStaffShift_(staffId);
      var dayStart = (shift.shiftStartMin != null ? shift.shiftStartMin : 540);
      var dayEnd   = (shift.shiftEndMin   != null ? shift.shiftEndMin   : 1080);

      // "隙間"を作成（アンカー前後にバッファを確保）
      var gaps = [];
      var cursor = dayStart;
      anchors.forEach(function(a){
        // アンカー開始前にバッファを空ける
        var gapEnd = a.s - EXTRA_BUFFER_MIN;
        if (cursor < gapEnd) gaps.push({ s: cursor, e: gapEnd });
        // アンカー終了後にバッファを空けて次へ
        cursor = Math.max(cursor, a.e + EXTRA_BUFFER_MIN);
      });
      if (cursor < dayEnd) gaps.push({ s: cursor, e: dayEnd });

      if (gaps.length === 0) return;

      // いったん flex を「所属gap」に振り分け（現状時刻で判定、取れないものは最初のgap）
      var flexByGap = gaps.map(function(){ return []; });

      flex.forEach(function(f){
        var placed = false;
        if (f.s != null && f.e != null) {
          for (var gi = 0; gi < gaps.length; gi++) {
            var g = gaps[gi];
            if (f.s >= g.s && f.e <= g.e) { flexByGap[gi].push(f); placed = true; break; }
          }
        }
        if (!placed) flexByGap[0].push(f);
      });

      // 各gap内で順序を最適化 → gap先頭から詰め直す
      for (var gi = 0; gi < gaps.length; gi++) {
        var g = gaps[gi];
        var list = flexByGap[gi];
        if (!list || list.length <= 1) continue;

        // llが無いものがあると距離最適化が弱いので、ll無しは末尾へ
        var withLL = list.filter(function(x){ return !!x.ll; });
        var noLL   = list.filter(function(x){ return !x.ll; });

        if (withLL.length >= 2) {
          var nn = nearestNeighbor_(withLL, staffLL);
          var opt = twoOpt_(nn);
          list = opt.concat(noLL);
        } else {
          list = withLL.concat(noLL);
        }

        // gapの先頭から詰め直し（バッファ込み）
        var t = g.s;
        for (var li = 0; li < list.length; li++) {
          var f = list[li];
          var s2 = t;
          var e2 = s2 + f.svcMin;

          if (e2 > g.e) {
            // 入らないなら触らない（安全優先）
            continue;
          }
          setRowTimeByMinutes_(f.row, s2, e2);
          t = e2 + EXTRA_BUFFER_MIN;
        }
      }
    });
  }

  // Level3距離最適化を実行
  applyLevel3RouteOptimizeSafe_();

  // ============================================================
  // 出力前に「時刻セル」をシリアル(0〜1)へ正規化（型混在対策）
  // ============================================================
  function normalizeTimeCellToSerial_(v) {
    var m = parseTimeToMinutes_(v);      // 既存のパーサを利用（Date/文字/数値すべて対応）
    if (m == null) return '';
    return minutesToSerial_(m);          // 必ずシリアルに統一
  }

  for (var i = 0; i < resultRows.length; i++) {
    var r = resultRows[i];

    // 開始/終了
    r[8] = normalizeTimeCellToSerial_(r[8]);
    r[9] = normalizeTimeCellToSerial_(r[9]);

    // 希望最早/希望最遅（空でもOK）
    r[12] = normalizeTimeCellToSerial_(r[12]);
    r[13] = normalizeTimeCellToSerial_(r[13]);
  }

  // 割当結果シートに書き込み
  resultSheet.clear();
  var header = ['visit_id','日付','曜日','staff_id','スタッフ名','patient_id','患者名','エリア',
                '開始時刻','終了時刻','サービス時間','時間タイプ','希望最早時刻','希望最遅時刻','備考',
                '前訪問ID','移動距離(km)','移動時間(分)'];
  resultSheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (resultRows.length > 0) {
    var outRows = resultRows.map(function(r, i){ return r.concat([prevVisitIdArr[i], moveKmArr[i], moveMinArr[i]]); });
    resultSheet.getRange(2, 1, outRows.length, header.length).setValues(outRows);
  }

  // 訪問履歴へ追加
  if (resultRows.length > 0) {
    var lastRow = historySheet.getLastRow();
    var histRows = resultRows.map(function(r){ return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[14]]; });
    historySheet.getRange(lastRow + 1, 1, histRows.length, 12).setValues(histRows);
  }

  // 割当不可シートへ出力
  var ngSheet = ss.getSheetByName('割当不可');
  if (!ngSheet) ngSheet = ss.insertSheet('割当不可');
  ngSheet.clear();
  ngSheet.getRange(1, 1, 1, 7).setValues([['日付', '曜日', 'patient_id', '患者名', '必要スタッフ数', '未割当枠', '理由']]);
  if (unassignedList.length > 0) {
    var ngOut = unassignedList.map(function(x){ return [x.date, x.youbi, x.pid, x.pname, x.needStaff, x.slot, x.reason]; });
    ngSheet.getRange(2, 1, ngOut.length, 7).setValues(ngOut);
  }

  return { message: '割当結果を ' + resultRows.length + ' 件作成しました。割当不可: ' + unassignedList.length + ' 件' };
}

// ============================================================
// 週間リクエストを生成（ss引数版）
// ============================================================

function 週間リクエストを生成_(ss) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('別の処理が実行中です。少し待ってから再実行してください。');

  try {
    const tz = ss.getSpreadsheetTimeZone();

    const patientSheet = ss.getSheetByName('患者マスタ');
    const weeklySheet  = ss.getSheetByName('週間リクエスト');
    const changeSheet  = ss.getSheetByName('個別変更リクエスト');
    const assignSheet  = ss.getSheetByName('割当結果');
    const historySheet = ss.getSheetByName('訪問履歴');

    if (!patientSheet || !changeSheet || !weeklySheet) {
      throw new Error('患者マスタ / 個別変更リクエスト / 週間リクエスト のシート名を確認してください。');
    }

    const pValues = patientSheet.getDataRange().getValues();
    if (pValues.length <= 1) throw new Error('「患者マスタ」にデータがありません。');
    const pHeader = pValues[0];
    const pData   = pValues.slice(1);

    const requiredHeaders = ['patient_id','患者名','エリア','週訪問回数','希望曜日（複数可）',
      '希望時間帯（開始）','希望時間帯（終了）','曜日NG','性別制限','継続希望','サービス時間','必要スタッフ数',
      '指定スタッフID','指定タイプ','NGスタッフID','備考'];

    const idx = {};
    const missing = [];
    requiredHeaders.forEach(h => { const i = pHeader.indexOf(h); if (i === -1) missing.push(h); else idx[h] = i; });
    if (missing.length > 0) throw new Error('患者マスタのヘッダーが足りません：\n' + missing.join('\n'));

    const timeTypeColIndex = pHeader.indexOf('時間タイプ');

    const patientInfoMap = {};
    pData.forEach(row => {
      const pid = row[idx['patient_id']];
      if (!pid) return;
      patientInfoMap[pid] = {
        name: row[idx['患者名']] || '', area: row[idx['エリア']] || '', svcMin: row[idx['サービス時間']],
        needStaff: toHalfWidthNumber_(row[idx['必要スタッフ数']], 1), sexLimit: row[idx['性別制限']],
        contPref: row[idx['継続希望']], timeType: (timeTypeColIndex >= 0 ? row[timeTypeColIndex] : '') || '',
        startPref: row[idx['希望時間帯（開始）']], endPref: row[idx['希望時間帯（終了）']],
        staffIds: row[idx['指定スタッフID']] || '', staffType: row[idx['指定タイプ']] || '',
        ngStaffIds: row[idx['NGスタッフID']] || '', note: row[idx['備考']] || ''
      };
    });

    const today = new Date();
    today.setHours(0,0,0,0);
    const day = today.getDay();
    const diffToMonday = (day + 6) % 7;
    const weekStart = new Date(today);
    weekStart.setDate(today.getDate() - diffToMonday);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    const weekStartStr = Utilities.formatDate(weekStart, tz, 'yyyy/MM/dd');
    const weekEndStr   = Utilities.formatDate(weekEnd, tz, 'yyyy/MM/dd');

    const lastVisitMap = {};
    if (historySheet) {
      const MAX_HISTORY_ROWS = 3000;
      const lastRow = historySheet.getLastRow();
      const lastCol = historySheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        const startRow = Math.max(2, lastRow - MAX_HISTORY_ROWS + 1);
        const numRows = lastRow - startRow + 1;
        const hHeader = historySheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const hData   = historySheet.getRange(startRow, 1, numRows, lastCol).getValues();
        const hIdxDate = hHeader.indexOf('日付'), hIdxPid = hHeader.indexOf('patient_id'),
              hIdxStaff = hHeader.indexOf('staff_id'), hIdxName = hHeader.indexOf('スタッフ名');
        hData.forEach(row => {
          const d = row[hIdxDate], pid = row[hIdxPid];
          if (!pid || !(d instanceof Date)) return;
          const ds = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
          if (ds >= weekStartStr) return;
          const current = lastVisitMap[pid];
          if (!current || d > current.date) lastVisitMap[pid] = { date: d, staffId: row[hIdxStaff] || '', staffName: row[hIdxName] || '' };
        });
      }
    }

    const youbiMap = { 'Sun': 0, 'Mon': 1, 'Tue': 2, 'Wed': 3, 'Thu': 4, 'Fri': 5, 'Sat': 6 };
    const indexToYoubi = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

    function parseDays(str) {
      if (!str) return [];
      const parts = String(str).split(/[,\u3001\/・\s]+/).map(s => s.trim()).filter(Boolean);
      const out = [];
      parts.forEach(p => { const y = normalizeYoubi(p); if (y && out.indexOf(y) === -1) out.push(y); });
      return out;
    }

    function toMinutes(v) {
      if (!v && v !== 0) return null;
      if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
      if (typeof v === 'number') return Math.round(v * 24 * 60);
      if (typeof v === 'string') { const parts = v.split(':'); if (parts.length >= 2) return Number(parts[0]) * 60 + Number(parts[1]); }
      return null;
    }

    function calcEndTime(startValue, minutes) {
      const m = Number(minutes || 0);
      if (!startValue || !m) return startValue;
      if (typeof startValue === 'number') return startValue + m / (24 * 60);
      else if (startValue instanceof Date) return new Date(startValue.getTime() + m * 60 * 1000);
      return startValue;
    }

    function makeTimeValue(h, m) { return (h * 60 + m) / (24 * 60); }

    function inferTimeType(timeTypeRaw, startPref, endPref, svcMin) {
      let t = (timeTypeRaw || '').trim();
      if (t) return t;
      if (!startPref || !endPref) return '固定';
      const s = toMinutes(startPref), e = toMinutes(endPref);
      if (s == null || e == null) return '固定';
      const span = e - s, svc = Number(svcMin || 0);
      if (!svc || span <= svc + 5) return '固定';
      if (s <= 9 * 60 + 15 && e >= 12 * 60 - 15) return '午前';
      if (s >= 13 * 60 - 15 && e <= 16 * 60 + 15) return '午後';
      if (span >= 7 * 60) return '終日';
      return '時間帯';
    }

    function makeTimeWindow(timeTypeRaw, startPref, endPref, svcMin) {
      const t = inferTimeType(timeTypeRaw, startPref, endPref, svcMin);
      let start = startPref, end = endPref, earliest = null, latest = null;
      if (t === '固定') { if (!end && start && svcMin) end = calcEndTime(start, svcMin); earliest = start; latest = end; }
      else if (t === '時間帯') { earliest = startPref; latest = endPref; if (!start && earliest) start = earliest; if (!end && start && svcMin) end = calcEndTime(start, svcMin); }
      else if (t === '午前') { earliest = makeTimeValue(9, 0); latest = makeTimeValue(12, 0); if (!start) start = earliest; if (!end && start && svcMin) end = calcEndTime(start, svcMin); }
      else if (t === '午後') { earliest = makeTimeValue(13, 0); latest = makeTimeValue(17, 0); if (!start) start = earliest; if (!end && start && svcMin) end = calcEndTime(start, svcMin); }
      else if (t === '終日') { earliest = makeTimeValue(9, 0); latest = makeTimeValue(18, 0); if (!start) start = earliest; if (!end && start && svcMin) end = calcEndTime(start, svcMin); }
      else { if (!end && start && svcMin) end = calcEndTime(start, svcMin); earliest = start; latest = end; }
      return { start, end, earliest, latest, timeType: t };
    }

    const weeklyRequests = [];
    pData.forEach(row => {
      const pid = row[idx['patient_id']];
      if (!pid) return;
      const info = patientInfoMap[pid] || {};
      const visits = toHalfWidthNumber_(row[idx['週訪問回数']], 0);
      if (!visits || visits <= 0) return;

      let prefDays = parseDays(row[idx['希望曜日（複数可）']]);
      const ngDays = parseDays(row[idx['曜日NG']]);
      if (prefDays.length === 0) prefDays = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
      const candidates = prefDays.filter(d => !ngDays.includes(d));
      if (candidates.length === 0) return;

      const actualVisits = Math.min(visits, candidates.length);
      const svcMin = info.svcMin, startPref = info.startPref, endPref = info.endPref, timeTypeRaw = info.timeType;
      const win = makeTimeWindow(timeTypeRaw, startPref, endPref, svcMin);
      const sexLimit = info.sexLimit, contPref = info.contPref, note = info.note || '';

      candidates.sort((a, b) => youbiMap[a] - youbiMap[b]);

      for (let i = 0; i < actualVisits; i++) {
        const youbi = candidates[i];
        const targetDay = youbiMap[youbi];
        for (let d = 0; d < 7; d++) {
          const dateObj = new Date(weekStart);
          dateObj.setDate(weekStart.getDate() + d);
          if (dateObj.getDay() === targetDay) {
            const dateStr = Utilities.formatDate(dateObj, tz, 'yyyy/MM/dd');
            const weekdayStr = Utilities.formatDate(dateObj, tz, 'EEE');
            const last = lastVisitMap[pid] || {};
            weeklyRequests.push({
              date: dateObj, dateStr: dateStr, weekdayStr: weekdayStr,
              patient_id: pid, patient_name: info.name || '', area: info.area || '',
              start: win.start, end: win.end, svcMin: svcMin, needStaff: info.needStaff || 1,
              specifiedIds: info.staffIds || '', specifiedType: info.staffType || '', ngStaffIds: info.ngStaffIds || '',
              sexLimit: sexLimit, contPref: contPref, changeType: '通常',
              prevStaffId: last.staffId || '', prevStaffName: last.staffName || '', prevDate: last.date || '',
              timeType: win.timeType, earliest: win.earliest, latest: win.latest, note: note
            });
            break;
          }
        }
      }
    });

    const weeklyMap = {};
    weeklyRequests.forEach(req => {
      const key = req.patient_id + '|' + req.dateStr;
      if (!weeklyMap[key]) weeklyMap[key] = [];
      weeklyMap[key].push(req);
    });

    // 個別変更リクエストの適用
    if (changeSheet) {
      const cValues = changeSheet.getDataRange().getValues();
      if (cValues.length > 1) {
        const cHeader = cValues[0], cData = cValues.slice(1);
        const cIdx = { patient_id: cHeader.indexOf('patient_id'), name: cHeader.indexOf('患者名'),
                       date: cHeader.indexOf('日付'), op: cHeader.indexOf('操作（キャンセル/時間変更/追加）'),
                       newStart: cHeader.indexOf('新開始時刻'), newEnd: cHeader.indexOf('新終了時刻'),
                       note: cHeader.indexOf('備考'), regAt: cHeader.indexOf('登録日時') };

        const changeMap = {};
        cData.forEach((row, idxRow) => {
          const pid = row[cIdx.patient_id], op = row[cIdx.op], d = row[cIdx.date];
          if (!pid || !op || !(d instanceof Date)) return;
          const dateStr = Utilities.formatDate(d, tz, 'yyyy/MM/dd');
          if (dateStr < weekStartStr || dateStr > weekEndStr) return;
          const key = pid + '|' + dateStr;
          let sortKey = cIdx.regAt !== -1 && row[cIdx.regAt] instanceof Date ? row[cIdx.regAt].getTime() : idxRow;
          const change = { pid, op, date: d, dateStr, newStart: row[cIdx.newStart], newEnd: row[cIdx.newEnd],
                          note: row[cIdx.note], patient_name: row[cIdx.name], sortKey };
          if (!changeMap[key] || sortKey > changeMap[key].sortKey) changeMap[key] = change;
        });

        Object.keys(changeMap).forEach(key => {
          const ch = changeMap[key];
          const matches = weeklyMap[ch.pid + '|' + ch.dateStr] || [];

          if (ch.op === 'キャンセル') {
            matches.forEach(req => { req.changeType = 'キャンセル'; if (ch.note) req.note = ch.note; });
          } else if (ch.op === '時間変更') {
            matches.forEach(req => {
              if (ch.newStart) req.start = ch.newStart;
              let endTime = ch.newEnd;
              if (!endTime && ch.newStart) endTime = calcEndTime(ch.newStart, req.svcMin);
              if (endTime) req.end = endTime;
              req.changeType = '変更';
              if (ch.note) req.note = ch.note;
            });
          } else if (ch.op === '追加') {
            const baseInfo = patientInfoMap[ch.pid] || {};
            const weekdayStr = indexToYoubi[ch.date.getDay()];
            let startTime = ch.newStart || baseInfo.startPref, endTime = ch.newEnd || baseInfo.endPref;
            const svcMin = baseInfo.svcMin || '', timeTypeRaw = baseInfo.timeType;
            const win = makeTimeWindow(timeTypeRaw, startTime, endTime, svcMin);
            const newReq = {
              date: ch.date, dateStr: ch.dateStr, weekdayStr: weekdayStr,
              patient_id: ch.pid, patient_name: ch.patient_name || baseInfo.name || '',
              area: baseInfo.area || '', start: win.start, end: win.end,
              svcMin: svcMin, needStaff: baseInfo.needStaff || 1,
              specifiedIds: baseInfo.staffIds || '', specifiedType: baseInfo.staffType || '', ngStaffIds: baseInfo.ngStaffIds || '',
              sexLimit: baseInfo.sexLimit || '', contPref: baseInfo.contPref || '',
              changeType: '追加', note: ch.note || baseInfo.note || '',
              prevStaffId: '', prevStaffName: '', prevDate: null,
              timeType: win.timeType, earliest: win.earliest, latest: win.latest
            };
            weeklyRequests.push(newReq);
            const mapKey = ch.pid + '|' + ch.dateStr;
            if (!weeklyMap[mapKey]) weeklyMap[mapKey] = [];
            weeklyMap[mapKey].push(newReq);
          }
        });
      }
    }

    // 割当結果から前回担当を付与
    const lastAssignMap = {};
    if (assignSheet) {
      const MAX_ASSIGN_ROWS = 2000;
      const lastRow = assignSheet.getLastRow(), lastCol = assignSheet.getLastColumn();
      if (lastRow > 1 && lastCol > 0) {
        const startRow = Math.max(2, lastRow - MAX_ASSIGN_ROWS + 1), numRows = lastRow - startRow + 1;
        const aHeader = assignSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const aData = assignSheet.getRange(startRow, 1, numRows, lastCol).getValues();
        const aIdxDate = aHeader.indexOf('日付'), aIdxPid = aHeader.indexOf('patient_id'),
              aIdxSid = aHeader.indexOf('staff_id'), aIdxSname = aHeader.indexOf('スタッフ名');
        aData.forEach(row => {
          const d = row[aIdxDate], pid = row[aIdxPid];
          if (!(d instanceof Date) || !pid) return;
          if (d >= weekStart) return;
          const dow = d.getDay(), key = pid + '|' + dow;
          const current = lastAssignMap[key];
          if (!current || d > current.date) lastAssignMap[key] = { staffId: row[aIdxSid], staffName: row[aIdxSname], date: d };
        });
      }
    }

    weeklyRequests.forEach(req => {
      const dow = req.date.getDay(), key = req.patient_id + '|' + dow;
      const last = lastAssignMap[key];
      if (last) { req.prevStaffId = last.staffId || ''; req.prevStaffName = last.staffName || ''; req.prevDate = last.date || null; }
    });

    if (weeklyRequests.length === 0) {
      throw new Error('週間リクエスト候補が0件でした。\n・「週訪問回数」が0または空ではないか\n・「希望曜日」と「曜日NG」の組み合わせで候補が消えていないか\n・今週の範囲（' + weekStartStr + '〜' + weekEndStr + '）でよいか\nなどを確認してください。');
    }

    weeklyRequests.sort((a, b) => {
      if (a.date - b.date !== 0) return a.date - b.date;
      const am = toMinutes(a.start), bm = toMinutes(b.start);
      if (am == null && bm == null) return 0;
      if (am == null) return 1;
      if (bm == null) return -1;
      return am - bm;
    });

    weeklySheet.clear();
    const headerOut = ['request_id','日付','曜日','patient_id','患者名','エリア',
      '開始時刻','終了時刻','サービス時間','必要スタッフ数','指定スタッフID','指定タイプ','NGスタッフID',
      '性別制限','継続希望','変更区分（通常/変更/追加/キャンセル）',
      '前回担当スタッフID','前回担当スタッフ名','前回訪問日','時間タイプ','希望最早時刻','希望最遅時刻','備考'];
    weeklySheet.getRange(1, 1, 1, headerOut.length).setValues([headerOut]);

    const out = weeklyRequests.map((req, i) => ([
      'R' + Utilities.formatString('%03d', i+1), req.date, req.weekdayStr, req.patient_id, req.patient_name, req.area,
      req.start, req.end, req.svcMin, req.needStaff || 1, req.specifiedIds || '', req.specifiedType || '', req.ngStaffIds || '',
      req.sexLimit, req.contPref, req.changeType, req.prevStaffId || '', req.prevStaffName || '', req.prevDate || '',
      req.timeType || '', req.earliest || '', req.latest || '', req.note
    ]));

    weeklySheet.getRange(2, 1, out.length, headerOut.length).setValues(out);

    return { message: '週間リクエストを ' + weekStartStr + ' 〜 ' + weekEndStr + ' 分、' + weeklyRequests.length + ' 件生成しました。' };

  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 位置情報を更新（ss引数版）
// ============================================================

function 位置情報を更新_(ss) {
  updateSheetLatLng_(ss.getSheetByName('患者マスタ'), '住所', '緯度', '経度');
  updateSheetLatLng_(ss.getSheetByName('スタッフマスタ'), '拠点住所', '緯度', '経度');
  return { message: '位置情報を更新しました' };
}

function updateSheetLatLng_(sheet, addrHeader, latHeader, lngHeader) {
  if (!sheet) return;
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return;
  const header = values[0], data = values.slice(1);
  const idxAddr = header.indexOf(addrHeader), idxLat = header.indexOf(latHeader), idxLng = header.indexOf(lngHeader);
  if (idxAddr === -1 || idxLat === -1 || idxLng === -1) return;

  const geocoder = Maps.newGeocoder();
  let changed = false;
  data.forEach((row, i) => {
    const addr = row[idxAddr], lat = row[idxLat], lng = row[idxLng];
    if (addr && (lat === '' || lng === '' || lat == null || lng == null)) {
      const res = geocoder.geocode(addr);
      if (res.status === 'OK' && res.results && res.results.length > 0) {
        const loc = res.results[0].geometry.location;
        data[i][idxLat] = loc.lat;
        data[i][idxLng] = loc.lng;
        changed = true;
      }
      Utilities.sleep(200);
    }
  });
  if (changed) sheet.getRange(2, 1, data.length, header.length).setValues(data);
}

// ============================================================
// ルートサマリを作成（ss引数版）
// ============================================================

function ルートサマリを作成_(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const resultSheet = ss.getSheetByName('割当結果');
  const patientSheet = ss.getSheetByName('患者マスタ');

  if (!resultSheet) throw new Error('「割当結果」シートが見つかりません。');
  if (!patientSheet) throw new Error('「患者マスタ」シートが見つかりません。');

  const pValues = patientSheet.getDataRange().getValues();
  const pHeader = pValues[0];
  const pIdx = { pid: pHeader.indexOf('patient_id'), addr: pHeader.indexOf('住所'),
                 lat: pHeader.indexOf('緯度'), lng: pHeader.indexOf('経度') };
  if (pIdx.pid === -1 || pIdx.addr === -1 || pIdx.lat === -1 || pIdx.lng === -1) {
    throw new Error('「患者マスタ」のヘッダー名（patient_id, 住所, 緯度, 経度）を確認してください。');
  }

  const patientMap = {};
  for (let i = 1; i < pValues.length; i++) {
    const row = pValues[i], pid = row[pIdx.pid];
    if (!pid) continue;
    patientMap[pid] = { addr: row[pIdx.addr] || '', lat: row[pIdx.lat] || '', lng: row[pIdx.lng] || '' };
  }

  const values = resultSheet.getDataRange().getValues();
  if (values.length <= 1) throw new Error('「割当結果」にデータがありません。');

  const header = values[0], data = values.slice(1);
  const idx = { date: header.indexOf('日付'), youbi: header.indexOf('曜日'), staffId: header.indexOf('staff_id'),
                sname: header.indexOf('スタッフ名'), dist: header.indexOf('移動距離(km)'), mtime: header.indexOf('移動時間(分)'),
                pid: header.indexOf('patient_id'), pname: header.indexOf('患者名'), start: header.indexOf('開始時刻') };

  if (idx.date === -1 || idx.staffId === -1 || idx.sname === -1 || idx.dist === -1 || idx.mtime === -1 ||
      idx.pid === -1 || idx.pname === -1 || idx.start === -1) {
    throw new Error('「割当結果」のヘッダー名を確認してください。');
  }

  const map = {};
  data.forEach(row => {
    const d = row[idx.date];
    if (!(d instanceof Date)) return;
    const staffId = row[idx.staffId], staffName = row[idx.sname];
    if (!staffId) return;
    const dateStr = Utilities.formatDate(d, tz, 'yyyy/MM/dd'), youbi = row[idx.youbi];
    const distKm = Number(row[idx.dist] || 0), moveMin = Number(row[idx.mtime] || 0);
    const key = staffId + '|' + dateStr;

    if (!map[key]) map[key] = { staffId, staffName, dateObj: d, dateStr, youbi,
                                visitCount: 0, moveCount: 0, distTotal: 0, timeTotal: 0, visits: [] };
    const rec = map[key];
    rec.visitCount++;
    if (distKm > 0 || moveMin > 0) { rec.moveCount++; rec.distTotal += distKm; rec.timeTotal += moveMin; }
    rec.visits.push({ start: row[idx.start], pid: row[idx.pid], pname: row[idx.pname] });
  });

  const records = Object.keys(map).map(k => map[k]);
  records.sort((a, b) => { if (a.staffName !== b.staffName) return a.staffName > b.staffName ? 1 : -1; return a.dateObj - b.dateObj; });

  let summarySheet = ss.getSheetByName('ルートサマリ');
  if (!summarySheet) summarySheet = ss.insertSheet('ルートサマリ');
  summarySheet.clear();

  const outHeader = ['staff_id','スタッフ名','日付','曜日','訪問件数','移動回数（前訪問あり）',
                     '総移動距離(km)','総移動時間(分)','ルート順（No. 患者ID 患者名 住所 (緯度, 経度)）'];
  summarySheet.getRange(1, 1, 1, outHeader.length).setValues([outHeader]);

  if (records.length > 0) {
    const out = records.map(r => {
      const visits = r.visits.slice().sort((a, b) => {
        if (!a.start && !b.start) return 0; if (!a.start) return 1; if (!b.start) return -1; return a.start - b.start;
      });
      const routeText = visits.map((v, idx) => {
        const p = patientMap[v.pid] || {};
        return 'No.' + (idx + 1) + ' ' + v.pid + ' ' + v.pname + ' ' + (p.addr || '') + ' (' + (p.lat || '') + ', ' + (p.lng || '') + ')';
      }).join(' → ');
      return [r.staffId, r.staffName, r.dateObj, r.youbi, r.visitCount, r.moveCount, r.distTotal, r.timeTotal, routeText];
    });
    summarySheet.getRange(2, 1, out.length, outHeader.length).setValues(out);
  }

  return { message: 'ルートサマリを ' + records.length + ' 行作成しました。' };
}


// ============================================================
// シート生成ユーティリティ
// ============================================================

/**
 * イベントリクエストシートを作成
 * GASエディタから直接実行可能
 */
function createEventRequestSheet() {
  var ssId = PropertiesService.getScriptProperties().getProperty("SS_ID");
  if (!ssId) {
    throw new Error("SS_ID が設定されていません");
  }
  var ss = SpreadsheetApp.openById(ssId);
  
  var sheetName = "イベントリクエスト";
  var existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    throw new Error("シート「" + sheetName + "」は既に存在します");
  }
  
  var sheet = ss.insertSheet(sheetName);
  
  var headers = [
    "event_id",
    "日付",
    "曜日",
    "staff_id",
    "スタッフ名",
    "イベント種別",
    "タイトル",
    "住所",
    "緯度",
    "経度",
    "時間指定方法",
    "開始時刻",
    "終了時刻",
    "所要時間(分)",
    "固定枠",
    "患者紐づき",
    "patient_id",
    "患者影響",
    "事前事務所戻り",
    "事後事務所戻り",
    "理由",
    "備考"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#4a7c59");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  
  // 列幅の調整
  sheet.setColumnWidth(1, 80);   // event_id
  sheet.setColumnWidth(2, 100);  // 日付
  sheet.setColumnWidth(3, 50);   // 曜日
  sheet.setColumnWidth(4, 80);   // staff_id
  sheet.setColumnWidth(5, 100);  // スタッフ名
  sheet.setColumnWidth(6, 120);  // イベント種別
  sheet.setColumnWidth(7, 150);  // タイトル
  sheet.setColumnWidth(8, 200);  // 住所
  sheet.setColumnWidth(11, 100); // 時間指定方法
  sheet.setColumnWidth(14, 100); // 所要時間(分)
  sheet.setColumnWidth(21, 200); // 理由
  sheet.setColumnWidth(22, 200); // 備考
  
  // 1行目を固定
  sheet.setFrozenRows(1);
  
  Logger.log("シート「" + sheetName + "」を作成しました");
  return "シート「" + sheetName + "」を作成しました";
}

