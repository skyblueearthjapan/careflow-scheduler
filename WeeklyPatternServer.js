/**
 * WeeklyPatternServer.js
 * GAS サーバー側ロジック: 週間訪問パターン管理
 *
 * Public functions (callable via google.script.run):
 *   - wp_getPatientList()
 *   - wp_getPattern(patientId)
 *   - wp_savePattern(patientId, patientName, slotsJson)
 *
 * Sheet: 週間訪問パターン
 *   patient_id | 患者名 | 曜日コード | 開始時刻 | 終了時刻 | サービス時間 | 必要スタッフ数 | 備考
 */

// ============================================================
// Public Functions
// ============================================================

/**
 * 患者マスタから患者一覧を取得（ドロップダウン用）
 * @returns {Object} { patients: [{pid, name, weeklyCount, needStaff, prefDays, ngDays,
 *                      timeType, startPref, endPref, svcMin, sexLimit, fixedStaff, contPref}] }
 */
function wp_getPatientList() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('患者マスタ');
    if (!sheet) {
      return { patients: [] };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { patients: [] };
    }

    var headers = data[0];
    var col = {};
    var colNames = [
      'patient_id', '患者名', '週訪問回数', '希望曜日（複数可）', '曜日NG',
      '必要スタッフ数', '性別制限', '継続希望', '指定スタッフID', '指定タイプ',
      'NGスタッフID', '時間タイプ', '希望時間帯（開始）', '希望時間帯（終了）',
      'サービス時間', '保険区分'
    ];
    for (var ci = 0; ci < colNames.length; ci++) {
      col[colNames[ci]] = headers.indexOf(colNames[ci]);
    }

    var patients = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var pid = col['patient_id'] >= 0 ? String(row[col['patient_id']]) : '';
      if (!pid) continue;

      var name = col['患者名'] >= 0 ? String(row[col['患者名']]) : '';
      var weeklyCount = col['週訪問回数'] >= 0 ? Number(row[col['週訪問回数']]) || 0 : 0;
      var needStaff = col['必要スタッフ数'] >= 0 ? Number(row[col['必要スタッフ数']]) || 1 : 1;
      var prefDays = col['希望曜日（複数可）'] >= 0 ? wp_parseDays_(row[col['希望曜日（複数可）']]) : [];
      var ngDays = col['曜日NG'] >= 0 ? wp_parseDays_(row[col['曜日NG']]) : [];
      var timeType = col['時間タイプ'] >= 0 ? String(row[col['時間タイプ']] || '') : '';
      var startPref = col['希望時間帯（開始）'] >= 0 ? wp_serialToMinutes_(row[col['希望時間帯（開始）']]) : null;
      var endPref = col['希望時間帯（終了）'] >= 0 ? wp_serialToMinutes_(row[col['希望時間帯（終了）']]) : null;
      var svcMin = col['サービス時間'] >= 0 ? Number(row[col['サービス時間']]) || 0 : 0;
      var sexLimit = col['性別制限'] >= 0 ? String(row[col['性別制限']] || '') : '';
      var fixedStaff = col['指定スタッフID'] >= 0 ? String(row[col['指定スタッフID']] || '') : '';
      var contPref = col['継続希望'] >= 0 ? String(row[col['継続希望']] || '') : '';

      patients.push({
        pid: pid,
        name: name,
        weeklyCount: weeklyCount,
        needStaff: needStaff,
        prefDays: prefDays,
        ngDays: ngDays,
        timeType: timeType,
        startPref: startPref,
        endPref: endPref,
        svcMin: svcMin,
        sexLimit: sexLimit,
        fixedStaff: fixedStaff,
        contPref: contPref
      });
    }

    // patient_id でソート
    patients.sort(function(a, b) {
      return a.pid < b.pid ? -1 : a.pid > b.pid ? 1 : 0;
    });

    return { patients: patients };
  } catch (e) {
    console.error('wp_getPatientList error: ' + e.message);
    throw new Error('患者一覧の取得に失敗しました: ' + e.message);
  }
}

/**
 * 指定患者の週間訪問パターンを取得
 * @param {string} patientId - 患者ID
 * @returns {Object} { slots: [{dayCode, startMin, endMin, svcMin, needStaff, note}] }
 */
function wp_getPattern(patientId) {
  if (!patientId || typeof patientId !== 'string') return { slots: [] };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('週間訪問パターン');
    if (!sheet) {
      return { slots: [] };
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return { slots: [] };
    }

    var headers = data[0];
    var colPid = headers.indexOf('patient_id');
    var colDay = headers.indexOf('曜日コード');
    var colStart = headers.indexOf('開始時刻');
    var colEnd = headers.indexOf('終了時刻');
    var colSvc = headers.indexOf('サービス時間');
    var colStaff = headers.indexOf('必要スタッフ数');
    var colNote = headers.indexOf('備考');

    var slots = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (colPid < 0 || String(row[colPid]) !== String(patientId)) continue;

      slots.push({
        dayCode: colDay >= 0 ? String(row[colDay]) : '',
        startMin: colStart >= 0 ? wp_serialToMinutes_(row[colStart]) : null,
        endMin: colEnd >= 0 ? wp_serialToMinutes_(row[colEnd]) : null,
        svcMin: colSvc >= 0 ? Number(row[colSvc]) || 0 : 0,
        needStaff: colStaff >= 0 ? Number(row[colStaff]) || 1 : 1,
        note: colNote >= 0 ? String(row[colNote] || '') : ''
      });
    }

    return { slots: slots };
  } catch (e) {
    console.error('wp_getPattern error: ' + e.message);
    throw new Error('パターン取得に失敗しました: ' + e.message);
  }
}

/**
 * 指定患者の週間訪問パターンを保存（既存行を削除してから挿入）
 * @param {string} patientId - 患者ID
 * @param {string} patientName - 患者名
 * @param {string} slotsJson - JSON文字列 [{dayCode, startMin, endMin, svcMin, needStaff, note}]
 * @returns {Object} { success: true, saved: numberOfSlots }
 */
function wp_savePattern(patientId, patientName, slotsJson) {
  try {
    var slots = JSON.parse(slotsJson);

    // Input validation and formula injection guard
    var VALID_DAY_CODES = {Mon:1,Tue:1,Wed:1,Thu:1,Fri:1,Sat:1,Sun:1};
    for (var s = 0; s < slots.length; s++) {
      var sl = slots[s];
      if (!VALID_DAY_CODES[sl.dayCode]) throw new Error('Invalid dayCode: ' + sl.dayCode);
      sl.startMin = Number(sl.startMin);
      sl.endMin = Number(sl.endMin);
      if (isNaN(sl.startMin) || sl.startMin < 0 || sl.startMin > 1439) throw new Error('Invalid startMin');
      if (isNaN(sl.endMin) || sl.endMin < 0 || sl.endMin > 1440) throw new Error('Invalid endMin');
      sl.note = String(sl.note || '').slice(0, 200);
      if (/^[=+\-@]/.test(sl.note)) sl.note = "'" + sl.note;
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = wp_ensureSheet_(ss);

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var colPid = headers.indexOf('patient_id');

      // 既存行を下から削除（インデックスずれ防止）
      if (colPid >= 0) {
        for (var i = data.length - 1; i >= 1; i--) {
          if (String(data[i][colPid]) === String(patientId)) {
            sheet.deleteRow(i + 1); // 1-indexed
          }
        }
      }

      // 新規行を追加
      for (var s = 0; s < slots.length; s++) {
        var slot = slots[s];
        var startSerial = wp_minutesToSerial_(slot.startMin);
        var endSerial = wp_minutesToSerial_(slot.endMin);
        var newRow = [
          patientId,
          patientName,
          slot.dayCode || '',
          startSerial,
          endSerial,
          Number(slot.svcMin) || 0,
          Number(slot.needStaff) || 1,
          slot.note || ''
        ];
        sheet.appendRow(newRow);
      }
    } finally {
      lock.releaseLock();
    }

    return { success: true, saved: slots.length };
  } catch (e) {
    console.error('wp_savePattern error: ' + e.message);
    throw new Error('パターン保存に失敗しました: ' + e.message);
  }
}

// ============================================================
// Internal Helper Functions
// ============================================================

/**
 * 週間訪問パターンシートの存在を確認し、無ければ作成する
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @returns {Sheet} 週間訪問パターンシート
 */
function wp_ensureSheet_(ss) {
  var sheet = ss.getSheetByName('週間訪問パターン');
  if (!sheet) {
    sheet = ss.insertSheet('週間訪問パターン');
    sheet.getRange(1, 1, 1, 8).setValues([[
      'patient_id', '患者名', '曜日コード', '開始時刻', '終了時刻',
      'サービス時間', '必要スタッフ数', '備考'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 分→時刻シリアル変換
 * @param {number} min - 0時からの分数
 * @returns {number} 時刻シリアル値
 */
function wp_minutesToSerial_(min) {
  if (min === null || min === undefined || isNaN(Number(min))) return null;
  return Number(min) / 60 / 24;
}

/**
 * 時刻シリアル/Date→分変換
 * @param {number|Date} serial - 時刻シリアル値またはDateオブジェクト
 * @returns {number|null} 0時からの分数 (0-1440)
 */
function wp_serialToMinutes_(serial) {
  if (serial instanceof Date) {
    return serial.getHours() * 60 + serial.getMinutes();
  }
  if (typeof serial === 'number') {
    return Math.round(serial * 24 * 60) % 1440;
  }
  return null;
}

/**
 * 曜日文字列をパースして配列にする
 * @param {string} str - カンマ/スペース区切りの曜日文字列
 * @returns {string[]} 曜日配列
 */
function wp_parseDays_(str) {
  if (!str) return [];
  return String(str).split(/[,、\s]+/).map(function(s) { return s.trim(); }).filter(Boolean);
}
