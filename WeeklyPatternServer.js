/**
 * WeeklyPatternServer.js
 * GAS サーバー側ロジック: 週間訪問パターン管理
 *
 * Public functions (callable via google.script.run):
 *   - openWeeklyPattern()
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
 * 週間訪問パターン画面のURLを返す（iframeモーダル用）
 * @returns {string} URL
 */
function openWeeklyPattern() {
  return ScriptApp.getService().getUrl() + '?page=pattern';
}

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
    // ヘッダー名の正規化（全角/半角スペース除去、小文字化）でマッチング
    var normalizedHeaders = [];
    for (var hi = 0; hi < headers.length; hi++) {
      normalizedHeaders.push(String(headers[hi] || '').replace(/[\s\u3000]/g, '').toLowerCase());
    }

    function findCol_(keyword) {
      // まず完全一致
      var exact = headers.indexOf(keyword);
      if (exact >= 0) return exact;
      // 正規化後の部分一致
      var kw = keyword.replace(/[\s\u3000]/g, '').toLowerCase();
      for (var fi = 0; fi < normalizedHeaders.length; fi++) {
        if (normalizedHeaders[fi].indexOf(kw) >= 0) return fi;
      }
      return -1;
    }

    var col = {
      pid:        findCol_('patient_id'),
      name:       findCol_('患者名'),
      weekly:     findCol_('週訪問回数'),
      prefDays:   findCol_('希望曜日'),
      ngDays:     findCol_('曜日NG'),
      needStaff:  findCol_('必要スタッフ数'),
      sexLimit:   findCol_('性別制限'),
      contPref:   findCol_('継続希望'),
      fixedStaff: findCol_('指定スタッフID'),
      fixedType:  findCol_('指定タイプ'),
      ngStaff:    findCol_('NGスタッフID'),
      timeType:   findCol_('時間タイプ'),
      startPref:  findCol_('希望時間帯（開始）'),
      endPref:    findCol_('希望時間帯（終了）'),
      svcMin:     findCol_('サービス時間'),
      insurance:  findCol_('保険区分'),
      status:     findCol_('稼働状況')
    };

    // デバッグ: 見つからなかった列をログ出力
    for (var ck in col) {
      if (col[ck] < 0) console.log('wp_getPatientList: column not found for key=' + ck);
    }

    var patients = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var pid = col.pid >= 0 ? String(row[col.pid] || '') : '';
      if (!pid) continue;

      var name = col.name >= 0 ? String(row[col.name] || '') : '';
      var weeklyCount = col.weekly >= 0 ? Number(row[col.weekly]) || 0 : 0;
      var needStaff = col.needStaff >= 0 ? Number(row[col.needStaff]) || 1 : 1;
      var prefDays = col.prefDays >= 0 ? wp_parseDays_(row[col.prefDays]) : [];
      var ngDays = col.ngDays >= 0 ? wp_parseDays_(row[col.ngDays]) : [];
      var timeType = col.timeType >= 0 ? String(row[col.timeType] || '') : '';
      var startPref = col.startPref >= 0 ? wp_serialToMinutes_(row[col.startPref]) : null;
      var endPref = col.endPref >= 0 ? wp_serialToMinutes_(row[col.endPref]) : null;
      var svcMin = col.svcMin >= 0 ? Number(row[col.svcMin]) || 0 : 0;
      var sexLimit = col.sexLimit >= 0 ? String(row[col.sexLimit] || '') : '';
      var fixedStaff = col.fixedStaff >= 0 ? String(row[col.fixedStaff] || '') : '';
      var contPref = col.contPref >= 0 ? String(row[col.contPref] || '') : '';

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
        contPref: contPref,
        status: col.status >= 0 ? String(row[col.status] || '') : ''
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
    var colTimeType = headers.indexOf('時間タイプ');

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
        note: colNote >= 0 ? String(row[colNote] || '') : '',
        timeType: colTimeType >= 0 ? String(row[colTimeType] || '') : ''
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
          slot.note || '',
          slot.timeType || ''
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
    sheet.getRange(1, 1, 1, 9).setValues([[
      'patient_id', '患者名', '曜日コード', '開始時刻', '終了時刻',
      'サービス時間', '必要スタッフ数', '備考', '時間タイプ'
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
