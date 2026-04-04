/**
 * 入力データ監査 - API層
 *
 * フロントエンド（UnifiedInput.html）から呼ばれるサーバーサイド関数。
 * google.script.run.inputAudit_runAudit(sheetName) の形で呼び出される。
 */

// ============================================================
// 対象シート一覧
// ============================================================

var INPUT_AUDIT_TARGET_SHEETS = [
  '患者マスタ',
  'スタッフマスタ',
  '個別変更リクエスト',
  'スタッフ個別変更リクエスト',
  'イベントリクエスト',
  'スタッフ同行割付'
];

// ============================================================
// 公開API
// ============================================================

/**
 * 指定シートの入力データを監査する
 * @param {string} sheetName - シート名
 * @return {Object} 監査結果
 */
function inputAudit_runAudit(sheetName) {
  try {
    var result = inputAudit_judgeSheet(sheetName);
    return {
      success: true,
      sheetName: sheetName,
      summary: result.summary,
      cellResults: result.cellResults,
      rowResults: result.rowResults,
      timestamp: new Date().toISOString()
    };
  } catch (e) {
    console.error('inputAudit_runAudit error: ' + e.message);
    return {
      success: false,
      error: e.message || String(e),
      sheetName: sheetName
    };
  }
}

/**
 * 全入力シートを一括監査する
 * @return {Object} { sheets, totalSummary }
 */
function inputAudit_runAllAudits() {
  var results = {};
  var totalNg = 0, totalWarn = 0, totalOk = 0;

  for (var i = 0; i < INPUT_AUDIT_TARGET_SHEETS.length; i++) {
    var sheetName = INPUT_AUDIT_TARGET_SHEETS[i];
    var result = inputAudit_runAudit(sheetName);
    results[sheetName] = result;
    if (result.success && result.summary) {
      totalNg += result.summary.ng;
      totalWarn += result.summary.warn;
      totalOk += result.summary.ok;
    }
  }

  return {
    success: true,
    sheets: results,
    totalSummary: { ng: totalNg, warn: totalWarn, ok: totalOk },
    timestamp: new Date().toISOString()
  };
}

/**
 * 特定セルの監査詳細を取得（モーダル表示用）
 * @param {string} sheetName
 * @param {number} rowIdx
 * @param {number} colIdx
 * @return {Object}
 */
function inputAudit_getCellDetail(sheetName, rowIdx, colIdx) {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty('SS_ID');
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'シートが見つかりません: ' + sheetName };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'データがありません' };

    var headers = data[0];
    var actualRowIdx = rowIdx + 1; // ヘッダー行分
    if (actualRowIdx >= data.length) return { success: false, error: '行が範囲外です' };

    var row = data[actualRowIdx];
    var fieldName = (colIdx >= 0 && colIdx < headers.length) ? String(headers[colIdx]).trim() : '';
    var cellValue = (colIdx >= 0 && colIdx < row.length) ? row[colIdx] : '';

    // 行全体を判定
    var issues = inputAudit_judgeRow_(sheetName, headers, row, rowIdx);

    // 該当セルのissueをフィルタ
    var cellIssues = [];
    var otherIssues = [];
    for (var i = 0; i < issues.length; i++) {
      if (issues[i].colIdx === colIdx) {
        cellIssues.push(issues[i]);
      } else {
        otherIssues.push(issues[i]);
      }
    }

    return {
      success: true,
      field: fieldName,
      value: cellValue !== null && cellValue !== undefined ? String(cellValue) : '',
      cellIssues: cellIssues,
      otherIssues: otherIssues,
      totalIssues: issues.length
    };
  } catch (e) {
    console.error('inputAudit_getCellDetail error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 特定行の全監査結果を取得（モーダル表示用）
 * @param {string} sheetName
 * @param {number} rowIdx
 * @return {Object}
 */
function inputAudit_getRowDetail(sheetName, rowIdx) {
  try {
    var ssId = PropertiesService.getScriptProperties().getProperty('SS_ID');
    var ss = SpreadsheetApp.openById(ssId);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'シートが見つかりません' };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, error: 'データがありません' };

    var headers = data[0];
    var actualRowIdx = rowIdx + 1;
    if (actualRowIdx >= data.length) return { success: false, error: '行が範囲外です' };

    var row = data[actualRowIdx];
    var issues = inputAudit_judgeRow_(sheetName, headers, row, rowIdx);

    // 行データをオブジェクトに変換
    var rowData = {};
    for (var i = 0; i < headers.length; i++) {
      var key = String(headers[i]).trim();
      if (key) {
        rowData[key] = (i < row.length && row[i] !== null && row[i] !== undefined) ? String(row[i]) : '';
      }
    }

    var ngCount = 0, warnCount = 0;
    for (var j = 0; j < issues.length; j++) {
      if (issues[j].level === 'NG') ngCount++;
      else if (issues[j].level === 'WARN') warnCount++;
    }

    return {
      success: true,
      rowData: rowData,
      issues: issues,
      summary: { ng: ngCount, warn: warnCount, ok: (ngCount === 0 && warnCount === 0) ? 1 : 0 },
      level: ngCount > 0 ? 'NG' : (warnCount > 0 ? 'WARN' : 'OK')
    };
  } catch (e) {
    console.error('inputAudit_getRowDetail error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * 各シートのNG/WARN件数を軽量に返す（タブバッジ表示用）
 * @return {Object}
 */
function inputAudit_getTabBadges() {
  var badges = {};
  for (var i = 0; i < INPUT_AUDIT_TARGET_SHEETS.length; i++) {
    var sheetName = INPUT_AUDIT_TARGET_SHEETS[i];
    try {
      var result = inputAudit_judgeSheet(sheetName);
      badges[sheetName] = { ng: result.summary.ng, warn: result.summary.warn };
    } catch (e) {
      badges[sheetName] = { ng: -1, warn: -1, error: e.message };
    }
  }
  return badges;
}
