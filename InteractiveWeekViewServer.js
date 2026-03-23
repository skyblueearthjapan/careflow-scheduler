/**
 * InteractiveWeekViewServer.js
 * GAS サーバー側ロジック: インタラクティブ週ビュー
 *
 * Public functions (callable via google.script.run):
 *   - openInteractiveWeekView(weekStartStr)
 *   - getInteractiveWeekData(weekStartStr)
 *   - commitChanges(weekStartStr, changesJson)
 *   - saveSnapshot(weekStartStr)
 *   - iwv_saveConfirmedHistory_(weekStartStr)
 *   - getPatientRotationHistory(patientId, weeksBack)
 */

// ============================================================
// Public Functions
// ============================================================

/**
 * インタラクティブ週ビューのURLを返す（Web App経由で開く）
 * @param {string} weekStartStr - "yyyy/MM/dd" 形式の週開始日
 * @return {string} インタラクティブ編集ページのURL
 */
function openInteractiveWeekView(weekStartStr) {
  var url = ScriptApp.getService().getUrl();
  url += '?page=interactive&week=' + encodeURIComponent(weekStartStr || '');
  return url;
}

/**
 * インタラクティブ週ビューに必要な全データを取得
 * @param {string} weekStartStr - "yyyy/MM/dd" 形式の週開始日
 * @returns {Object} 全データを含むオブジェクト
 */
function getInteractiveWeekData(weekStartStr) {
  var ss = iwv_getSpreadsheet_();

  var weekStart = iwv_parseDate_(weekStartStr);
  if (!weekStart) throw new Error('無効な週開始日: ' + weekStartStr);

  var weekDates = [];
  for (var d = 0; d < 7; d++) {
    var dt = new Date(weekStart);
    dt.setDate(dt.getDate() + d);
    weekDates.push(iwv_formatDate_(dt));
  }

  // 割当結果から全行を読み込み、割当済み / 未割当を分離
  var allRows = iwv_loadAssignments_(ss, weekDates);
  var assigned = [];
  var unassigned = [];
  for (var ai = 0; ai < allRows.length; ai++) {
    var row = allRows[ai];
    if (!row.staffId || row.staffName === '未割当') {
      row.staffId = '__UNASSIGNED__';
      unassigned.push(row);
    } else {
      assigned.push(row);
    }
  }
  // 割当不可シートのみにいる患者を補完（割当結果に無いケース）
  var ngList = iwv_loadUnassigned_(ss, weekDates);
  var unassignedPidDate = {};
  for (var ui = 0; ui < unassigned.length; ui++) {
    unassignedPidDate[unassigned[ui].pid + '|' + unassigned[ui].dateStr] = true;
  }
  for (var ni = 0; ni < ngList.length; ni++) {
    var ng = ngList[ni];
    if (!unassignedPidDate[ng.pid + '|' + ng.dateStr]) {
      ng.staffId = '__UNASSIGNED__';
      ng.visitId = ng.visitId || ('UA_' + ng.pid + '_' + ng.dateStr);
      unassigned.push(ng);
    }
  }

  // 2名体制の不足分を未割当に追加
  var patientMap = iwv_loadPatientMaster_(ss);
  var pidDateAssigned = {};
  var pidDateRef = {};
  for (var asi = 0; asi < assigned.length; asi++) {
    var ar = assigned[asi];
    if (!ar.pid || !ar.dateStr) continue;
    var pdKey = ar.pid + '|' + ar.dateStr;
    pidDateAssigned[pdKey] = (pidDateAssigned[pdKey] || 0) + 1;
    if (!pidDateRef[pdKey]) pidDateRef[pdKey] = ar;
  }
  for (var pdKey in pidDateAssigned) {
    var parts = pdKey.split('|');
    var pid = parts[0];
    var dateStr = parts[1];
    var pat = patientMap[pid];
    if (!pat || !pat.needStaff || pat.needStaff < 2) continue;
    var actual = pidDateAssigned[pdKey];
    if (actual < pat.needStaff) {
      var ref = pidDateRef[pdKey];
      var missing = pat.needStaff - actual;
      for (var mi = 0; mi < missing; mi++) {
        unassigned.push({
          visitId: (ref.visitId || '') + '_NEED' + mi,
          dateStr: dateStr,
          weekday: ref.weekday || '',
          pid: pid,
          pname: ref.pname || (pat.name || ''),
          staffId: '__UNASSIGNED__',
          staffName: '',
          startMin: ref.startMin || null,
          endMin: ref.endMin || null,
          svcMin: ref.svcMin || (pat.svcMin || 30),
          timeType: ref.timeType || '',
          needStaff: pat.needStaff,
          area: ref.area || '',
          note: '[2名体制:2人目未割当]'
        });
      }
    }
  }

  return {
    weekStartStr: weekStartStr,
    weekDates: weekDates,
    staffList: iwv_loadStaffMaster_(ss),
    patientMap: patientMap,
    assignments: assigned,
    unassigned: unassigned,
    eventMap: iwv_loadEvents_(ss, weekDates),
    staffChangeMap: iwv_loadStaffChanges_(ss, weekDates),
    changeRequests: iwv_loadChangeRequests_(ss, weekDates),
    specialWeek: iwv_loadSpecialWeek_(ss, weekDates)
  };
}

/**
 * 手動変更を割当結果シートに反映
 * @param {string} weekStartStr - 週開始日
 * @param {string} changesJson - 変更配列のJSON文字列
 * @returns {Object} 結果
 */
function commitChanges(weekStartStr, changesJson) {
  var changes = JSON.parse(changesJson);
  if (!changes || changes.length === 0) {
    return { success: true, changesApplied: 0 };
  }

  var ss = iwv_getSpreadsheet_();
  var sheet = ss.getSheetByName('割当結果');
  if (!sheet) throw new Error('割当結果シートが見つかりません');

  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var col = {
    visitId:   iwv_findHeaderIndex_(headers, 'visit_id'),
    staffId:   iwv_findHeaderIndex_(headers, 'staff_id'),
    staffName: iwv_findHeaderIndex_(headers, 'スタッフ名'),
    startTime: iwv_findHeaderIndex_(headers, '開始時刻'),
    endTime:   iwv_findHeaderIndex_(headers, '終了時刻'),
    note:      iwv_findHeaderIndex_(headers, '備考')
  };

  // patient_id / 日付カラム（ログ用）
  var colPid  = iwv_findHeaderIndex_(headers, 'patient_id');
  var colDate = iwv_findHeaderIndex_(headers, '日付');
  var colPname = iwv_findHeaderIndex_(headers, '患者名');
  var colYoubi   = iwv_findHeaderIndex_(headers, '曜日');
  var colArea    = iwv_findHeaderIndex_(headers, 'エリア');
  var colSvcMin  = iwv_findHeaderIndex_(headers, 'サービス時間');
  var colTType   = iwv_findHeaderIndex_(headers, '時間タイプ');

  var applied = 0;
  var logEntries = [];
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  changes.forEach(function(change) {
    var found = false;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][col.visitId]) === String(change.visitId)) {
        found = true;

        // --- 変更前の値をキャプチャ ---
        var beforeStaffId   = String(data[r][col.staffId] || '');
        var beforeStaffName = String(data[r][col.staffName] || '');
        var beforeStartMin  = iwv_serialToMinutes_(data[r][col.startTime]);
        var beforeEndMin    = iwv_serialToMinutes_(data[r][col.endTime]);

        // スタッフ変更
        if (change.newStaffId !== undefined && change.newStaffId !== null) {
          data[r][col.staffId] = change.newStaffId;
          data[r][col.staffName] = change.newStaffName || '';
        }
        // 時間変更 (minutes → serial) - 開始・終了の両方が必要
        if (change.newStartMin != null && change.newEndMin != null) {
          data[r][col.startTime] = iwv_minutesToSerial_(change.newStartMin);
          data[r][col.endTime] = iwv_minutesToSerial_(change.newEndMin);
        }
        // 日付変更
        if (change.newDateStr) {
          var newDate = iwv_parseDate_(change.newDateStr);
          if (colDate >= 0) {
            data[r][colDate] = newDate || change.newDateStr;
          }
          if (colYoubi >= 0 && newDate) {
            data[r][colYoubi] = iwv_getYoubiJa_(newDate);
          }
        }
        // 備考に手動変更マーク追加
        var note = String(data[r][col.note] || '');
        if (note.indexOf('[手動変更]') === -1) {
          if (note && !note.endsWith(' ')) note += ' ';
          note += '[手動変更]';
          data[r][col.note] = note;
        }

        // --- 変更種別を判定 ---
        var staffChanged = (change.newStaffId !== undefined && change.newStaffId !== null
                            && String(change.newStaffId) !== beforeStaffId);
        var timeChanged  = (change.newStartMin != null && change.newEndMin != null
                            && (change.newStartMin !== beforeStartMin || change.newEndMin !== beforeEndMin));
        // Bug5 fix: compare newDateStr against origDateStr (original date), falling back to dateStr
        var dateChanged  = (change.newDateStr && change.newDateStr !== (change.origDateStr || change.dateStr));
        var changeType = '';
        if (dateChanged && staffChanged && timeChanged) changeType = 'スタッフ+時間+曜日変更';
        else if (dateChanged && staffChanged)           changeType = 'スタッフ+曜日変更';
        else if (dateChanged && timeChanged)            changeType = '時間+曜日変更';
        else if (dateChanged)                           changeType = '曜日変更';
        else if (staffChanged && timeChanged)           changeType = 'スタッフ+時間変更';
        else if (staffChanged)                          changeType = 'スタッフ変更';
        else if (timeChanged)                           changeType = '時間変更';
        else                                            changeType = '更新（差分なし）';

        // --- ログ用の患者情報 ---
        var pid   = change.pid   || (colPid >= 0 ? String(data[r][colPid] || '') : '');
        var pname = change.pname || (colPname >= 0 ? String(data[r][colPname] || '') : '');
        var dateStr = change.dateStr || '';
        if (!dateStr && colDate >= 0) {
          var dv = data[r][colDate];
          if (dv instanceof Date) dateStr = iwv_formatDate_(dv);
          else if (dv) { var pd = iwv_parseDate_(dv); dateStr = pd ? iwv_formatDate_(pd) : String(dv); }
        }

        logEntries.push([
          now,                                                 // 確定日時
          String(change.visitId),                              // visit_id
          pid,                                                 // patient_id
          pname,                                               // 患者名
          dateStr,                                             // 日付
          changeType,                                          // 変更種別
          beforeStaffId,                                       // 変更前staff_id
          beforeStaffName,                                     // 変更前スタッフ名
          String(change.newStaffId != null ? change.newStaffId : beforeStaffId),   // 変更後staff_id
          change.newStaffName || (change.newStaffId != null ? '' : beforeStaffName), // 変更後スタッフ名
          beforeStartMin != null ? iwv_minutesToHHMM_(beforeStartMin) : '',        // 変更前開始
          beforeEndMin   != null ? iwv_minutesToHHMM_(beforeEndMin)   : '',        // 変更前終了
          change.newStartMin != null ? iwv_minutesToHHMM_(change.newStartMin) : (beforeStartMin != null ? iwv_minutesToHHMM_(beforeStartMin) : ''), // 変更後開始
          change.newEndMin   != null ? iwv_minutesToHHMM_(change.newEndMin)   : (beforeEndMin   != null ? iwv_minutesToHHMM_(beforeEndMin)   : ''), // 変更後終了
          weekStartStr                                         // 週開始日
        ]);

        applied++;
        break;
      }
    }
    if (!found) {
      // UA_ prefix の visitId → 割当結果シートに新規行を INSERT
      if (String(change.visitId).indexOf('UA_') === 0 &&
          change.newStaffId && change.pid && change.dateStr) {
        var newRow = [];
        for (var ci = 0; ci < headers.length; ci++) newRow.push('');

        // 決定論的な新visitId: pid_dateStr(スラッシュ除去)_staffId
        var newVisitId = change.pid + '_' + String(change.dateStr).replace(/\//g, '') + '_' + change.newStaffId;
        if (col.visitId >= 0)   newRow[col.visitId]   = newVisitId;
        if (colPid >= 0)        newRow[colPid]         = change.pid;
        if (colPname >= 0)      newRow[colPname]       = change.pname || '';
        if (col.staffId >= 0)   newRow[col.staffId]    = change.newStaffId;
        if (col.staffName >= 0) newRow[col.staffName]  = change.newStaffName || '';
        if (colDate >= 0) {
          var dp = iwv_parseDate_(change.dateStr);
          newRow[colDate] = dp || change.dateStr;
        }
        if (colYoubi >= 0 && change.dateStr) {
          var dp2 = iwv_parseDate_(change.dateStr);
          newRow[colYoubi] = dp2 ? iwv_getYoubiJa_(dp2) : '';
        }
        if (col.startTime >= 0 && change.newStartMin != null) {
          newRow[col.startTime] = iwv_minutesToSerial_(change.newStartMin);
        }
        if (col.endTime >= 0 && change.newEndMin != null) {
          newRow[col.endTime] = iwv_minutesToSerial_(change.newEndMin);
        }
        if (colArea >= 0)    newRow[colArea]    = change.area || '';
        if (colSvcMin >= 0)  newRow[colSvcMin]  = change.svcMin || '';
        if (colTType >= 0)   newRow[colTType]   = change.timeType || '';
        if (col.note >= 0)   newRow[col.note]   = '[手動挿入]';

        data.push(newRow);

        var pid   = change.pid || '';
        var pname = change.pname || '';
        var dateStr = change.dateStr || '';
        logEntries.push([
          now, newVisitId, pid, pname, dateStr,
          '新規挿入',
          '', '',  // 変更前staff (なし)
          String(change.newStaffId), change.newStaffName || '',
          '', '',  // 変更前時間 (なし)
          change.newStartMin != null ? iwv_minutesToHHMM_(change.newStartMin) : '',
          change.newEndMin   != null ? iwv_minutesToHHMM_(change.newEndMin)   : '',
          weekStartStr
        ]);
        applied++;
      } else {
        Logger.log('commitChanges: visit_id "' + change.visitId + '" が割当結果に見つかりません');
      }
    }
  });

  // シートに書き戻し
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  // --- 手動変更ログシートに追記 ---
  if (logEntries.length > 0) {
    iwv_appendChangeLog_(ss, logEntries);
  }

  // 依存シートを再構築
  try {
    割当不可を再構築_(ss);
  } catch (e) {
    Logger.log('割当不可を再構築_ エラー: ' + e.message);
  }
  try {
    週ビューを更新_(ss, weekStartStr);
  } catch (e) {
    Logger.log('週ビューを更新_ エラー: ' + e.message);
  }

  return { success: true, changesApplied: applied };
}

/**
 * 割当結果と週ビューのスナップショットを保存
 * @param {string} weekStartStr - 週開始日
 * @returns {Object} 結果 {success, resultSheetName, weekSheetName, label}
 */
function saveSnapshot(weekStartStr) {
  var ss = iwv_getSpreadsheet_();
  var weekStart = iwv_parseDate_(weekStartStr);
  if (!weekStart) throw new Error('無効な週開始日: ' + weekStartStr);

  var year = weekStart.getFullYear();
  var month = weekStart.getMonth() + 1;
  var weekOfMonth = iwv_getWeekOfMonth_(weekStart);
  var label = year + '年' + month + '月' + weekOfMonth + '週目';
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MMdd_HHmm');

  // 割当結果コピー
  var resultSheet = ss.getSheetByName('割当結果');
  var resultName = '割当結果_' + label;
  if (ss.getSheetByName(resultName)) {
    resultName = '割当結果_' + label + '_' + timestamp;
  }
  if (resultSheet) {
    var resultCopy = resultSheet.copyTo(ss);
    resultCopy.setName(resultName);
  }

  // 週ビューコピー
  var weekSheet = ss.getSheetByName('週ビュー');
  var weekName = '週ビュー_' + label;
  if (ss.getSheetByName(weekName)) {
    weekName = '週ビュー_' + label + '_' + timestamp;
  }
  if (weekSheet) {
    var weekCopy = weekSheet.copyTo(ss);
    weekCopy.setName(weekName);
  }

  // 確定スケジュール履歴に保存（スナップショットコピーから読む）
  var historySaved = 0;
  try {
    var histResult = iwv_saveConfirmedHistory_(weekStartStr, resultName);
    historySaved = histResult.saved;
    Logger.log('確定スケジュール履歴: ' + historySaved + '件保存 (from ' + resultName + ')');
  } catch (e) {
    Logger.log('確定スケジュール履歴エラー: ' + e.message);
  }

  return {
    success: true,
    resultSheetName: resultName,
    weekSheetName: weekName,
    label: label,
    historySaved: historySaved
  };
}

// ============================================================
// Private Helper Functions (iwv_ prefix)
// ============================================================

/**
 * スプレッドシートを取得（Web Appコンテキスト対応）
 */
function iwv_getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  if (!ss) throw new Error('スプレッドシートが見つかりません');
  return ss;
}

/**
 * 日付文字列をDateオブジェクトに変換
 */
function iwv_parseDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  var s = String(val).trim();
  // "yyyy/MM/dd" or "yyyy-MM-dd"
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

/**
 * Dateオブジェクトを "yyyy/MM/dd" に変換
 */
function iwv_formatDate_(date) {
  if (!date) return '';
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '/' + m + '/' + d;
}

/**
 * 英語曜日略称を取得
 */
function iwv_getYoubiEn_(date) {
  var days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  return days[date.getDay()];
}

/**
 * 日本語曜日を取得
 */
function iwv_getYoubiJa_(date) {
  return ['日','月','火','水','木','金','土'][date.getDay()];
}

/**
 * 曜日を英語略称に正規化
 */
function iwv_normalizeYoubi_(youbi) {
  if (!youbi) return '';
  var s = String(youbi).trim();
  var map = {
    '月': 'Mon', '火': 'Tue', '水': 'Wed', '木': 'Thu',
    '金': 'Fri', '土': 'Sat', '日': 'Sun',
    'Mon': 'Mon', 'Tue': 'Tue', 'Wed': 'Wed', 'Thu': 'Thu',
    'Fri': 'Fri', 'Sat': 'Sat', 'Sun': 'Sun',
    '月曜': 'Mon', '火曜': 'Tue', '水曜': 'Wed', '木曜': 'Thu',
    '金曜': 'Fri', '土曜': 'Sat', '日曜': 'Sun'
  };
  return map[s] || s;
}

/**
 * 時間値を分に変換（Date, serial, "HH:mm" 対応）
 */
function iwv_serialToMinutes_(val) {
  if (val === null || val === undefined || val === '') return null;
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }
  if (typeof val === 'number') {
    if (val >= 0 && val < 1.5) {
      // Serial (0.375 = 09:00 = 540分)
      return Math.round(val * 1440);
    }
    // Already minutes
    return Math.round(val);
  }
  if (typeof val === 'string') {
    var s = val.trim();
    // "HH:mm" or "H:mm"
    var m = s.match(/^(\d{1,2}):(\d{2})/);
    if (m) return parseInt(m[1]) * 60 + parseInt(m[2]);
    // "HH時mm分"
    m = s.match(/(\d{1,2})時(\d{1,2})分/);
    if (m) return parseInt(m[1]) * 60 + parseInt(m[2]);
  }
  return null;
}

/**
 * 分をスプレッドシートのシリアル値に変換
 */
function iwv_minutesToSerial_(min) {
  if (min === null || min === undefined) return '';
  return min / 1440;
}

/**
 * 月内の週番号を計算（月曜基準, 1-based）
 */
function iwv_getWeekOfMonth_(date) {
  var firstDay = new Date(date.getFullYear(), date.getMonth(), 1);
  var firstMonday = new Date(firstDay);
  while (firstMonday.getDay() !== 1) {
    firstMonday.setDate(firstMonday.getDate() + 1);
  }
  if (date < firstMonday) return 1;
  return Math.floor((date - firstMonday) / (7 * 24 * 60 * 60 * 1000)) + 1;
}

/**
 * 曜日文字列をパースして英語略称配列に変換
 * "月,水,金" → ["Mon","Wed","Fri"]
 */
function iwv_parseDays_(daysStr) {
  if (!daysStr) return [];
  return String(daysStr).split(/[,、\s]+/).map(function(d) {
    return iwv_normalizeYoubi_(d.trim());
  }).filter(function(d) { return d; });
}

/**
 * ヘッダー行からカラムインデックスを検索（柔軟マッチング）
 */
function iwv_findHeaderIndex_(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (h === name) return i;
  }
  // 部分一致フォールバック
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (h.indexOf(name) >= 0 || name.indexOf(h) >= 0) return i;
  }
  // 全角半角変換して再試行
  var nameHalf = name.replace(/[\uff01-\uff5e]/g, function(c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().replace(/[\uff01-\uff5e]/g, function(c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    });
    if (h === nameHalf) return i;
  }
  return -1;
}

// ============================================================
// Data Loaders
// ============================================================

/**
 * スタッフマスタを読み込み
 */
function iwv_loadStaffMaster_(ss) {
  var sheet = ss.getSheetByName('スタッフマスタ');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = data[0];

  var ci = {
    id:        iwv_findHeaderIndex_(h, 'staff_id'),
    name:      iwv_findHeaderIndex_(h, 'スタッフ名'),
    gender:    iwv_findHeaderIndex_(h, '性別'),
    shiftStart:iwv_findHeaderIndex_(h, 'シフト開始'),
    shiftEnd:  iwv_findHeaderIndex_(h, 'シフト終了'),
    workDays:  iwv_findHeaderIndex_(h, '勤務曜日'),
    maxPerDay: iwv_findHeaderIndex_(h, '最大訪問件数/日'),
    allocPref: iwv_findHeaderIndex_(h, '割当量'),
    areas:     iwv_findHeaderIndex_(h, '得意エリア'),
    lat:       iwv_findHeaderIndex_(h, '緯度'),
    lng:       iwv_findHeaderIndex_(h, '経度')
  };

  var result = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var id = ci.id >= 0 ? String(row[ci.id]).trim() : '';
    if (!id) continue;
    result.push({
      id: id,
      name: ci.name >= 0 ? String(row[ci.name]).trim() : '',
      gender: ci.gender >= 0 ? String(row[ci.gender]).trim() : '',
      shiftStartMin: ci.shiftStart >= 0 ? (iwv_serialToMinutes_(row[ci.shiftStart]) || 540) : 540,
      shiftEndMin: ci.shiftEnd >= 0 ? (iwv_serialToMinutes_(row[ci.shiftEnd]) || 1020) : 1020,
      workDays: ci.workDays >= 0 ? iwv_parseDays_(row[ci.workDays]) : ['Mon','Tue','Wed','Thu','Fri'],
      maxPerDay: ci.maxPerDay >= 0 ? (parseInt(row[ci.maxPerDay]) || 999) : 999,
      allocPref: ci.allocPref >= 0 ? String(row[ci.allocPref]).trim() : '均等',
      areas: ci.areas >= 0 ? String(row[ci.areas]).split(/[,、\s]+/).filter(Boolean) : [],
      lat: ci.lat >= 0 ? Number(row[ci.lat]) || null : null,
      lng: ci.lng >= 0 ? Number(row[ci.lng]) || null : null
    });
  }
  return result;
}

/**
 * 患者マスタを読み込み
 */
function iwv_loadPatientMaster_(ss) {
  var sheet = ss.getSheetByName('患者マスタ');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var h = data[0];

  var ci = {
    pid:        iwv_findHeaderIndex_(h, 'patient_id'),
    name:       iwv_findHeaderIndex_(h, '患者名'),
    area:       iwv_findHeaderIndex_(h, 'エリア'),
    sexLimit:   iwv_findHeaderIndex_(h, '性別制限'),
    needStaff:  iwv_findHeaderIndex_(h, '必要スタッフ数'),
    fixedStaff: iwv_findHeaderIndex_(h, '指定スタッフID'),
    fixedType:  iwv_findHeaderIndex_(h, '指定タイプ'),
    ngStaff:    iwv_findHeaderIndex_(h, 'NGスタッフID'),
    contPref:   iwv_findHeaderIndex_(h, '継続希望'),
    timeType:   iwv_findHeaderIndex_(h, '時間タイプ'),
    svcMin:     iwv_findHeaderIndex_(h, 'サービス時間'),
    prefDays:   iwv_findHeaderIndex_(h, '希望曜日'),
    ngDays:     iwv_findHeaderIndex_(h, '曜日NG'),
    startPref:  iwv_findHeaderIndex_(h, '希望時間帯（開始）'),
    endPref:    iwv_findHeaderIndex_(h, '希望時間帯（終了）'),
    weeklyCount:iwv_findHeaderIndex_(h, '週訪問回数'),
    insurance:  iwv_findHeaderIndex_(h, '保険区分'),
    lat:        iwv_findHeaderIndex_(h, '緯度'),
    lng:        iwv_findHeaderIndex_(h, '経度')
  };

  var map = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var pid = ci.pid >= 0 ? String(row[ci.pid]).trim() : '';
    if (!pid) continue;
    map[pid] = {
      pid: pid,
      name: ci.name >= 0 ? String(row[ci.name]).trim() : '',
      area: ci.area >= 0 ? String(row[ci.area]).trim() : '',
      sexLimit: ci.sexLimit >= 0 ? String(row[ci.sexLimit]).trim() : '',
      needStaff: ci.needStaff >= 0 ? (parseInt(String(row[ci.needStaff]).replace(/[^\d]/g, '')) || 1) : 1,
      fixedStaff: ci.fixedStaff >= 0 ? String(row[ci.fixedStaff]).trim() : '',
      fixedType: ci.fixedType >= 0 ? String(row[ci.fixedType]).trim() : '',
      ngStaff: ci.ngStaff >= 0 ? String(row[ci.ngStaff]).trim() : '',
      contPref: ci.contPref >= 0 ? String(row[ci.contPref]).trim() : '',
      timeType: ci.timeType >= 0 ? String(row[ci.timeType]).trim() : '',
      svcMin: ci.svcMin >= 0 ? (parseInt(row[ci.svcMin]) || 30) : 30,
      prefDays: ci.prefDays >= 0 ? iwv_parseDays_(row[ci.prefDays]) : [],
      ngDays: ci.ngDays >= 0 ? iwv_parseDays_(row[ci.ngDays]) : [],
      startPref: ci.startPref >= 0 ? iwv_serialToMinutes_(row[ci.startPref]) : null,
      endPref: ci.endPref >= 0 ? iwv_serialToMinutes_(row[ci.endPref]) : null,
      weeklyCount: ci.weeklyCount >= 0 ? (parseInt(row[ci.weeklyCount]) || 0) : 0,
      insuranceType: ci.insurance >= 0 ? String(row[ci.insurance]).trim() : '',
      lat: ci.lat >= 0 ? Number(row[ci.lat]) || null : null,
      lng: ci.lng >= 0 ? Number(row[ci.lng]) || null : null
    };
  }
  return map;
}

/**
 * 割当結果を読み込み（週でフィルタ）
 */
function iwv_loadAssignments_(ss, weekDates) {
  var sheet = ss.getSheetByName('割当結果');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = data[0];

  var ci = {
    visitId:   iwv_findHeaderIndex_(h, 'visit_id'),
    date:      iwv_findHeaderIndex_(h, '日付'),
    youbi:     iwv_findHeaderIndex_(h, '曜日'),
    staffId:   iwv_findHeaderIndex_(h, 'staff_id'),
    staffName: iwv_findHeaderIndex_(h, 'スタッフ名'),
    pid:       iwv_findHeaderIndex_(h, 'patient_id'),
    pname:     iwv_findHeaderIndex_(h, '患者名'),
    area:      iwv_findHeaderIndex_(h, 'エリア'),
    start:     iwv_findHeaderIndex_(h, '開始時刻'),
    end:       iwv_findHeaderIndex_(h, '終了時刻'),
    svcMin:    iwv_findHeaderIndex_(h, 'サービス時間'),
    timeType:  iwv_findHeaderIndex_(h, '時間タイプ'),
    earliest:  iwv_findHeaderIndex_(h, '希望最早時刻'),
    latest:    iwv_findHeaderIndex_(h, '希望最遅時刻'),
    note:      iwv_findHeaderIndex_(h, '備考')
  };

  var dateSet = {};
  weekDates.forEach(function(d) { dateSet[d] = true; });

  var result = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = ci.date >= 0 ? row[ci.date] : null;
    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = iwv_formatDate_(dateVal);
    } else if (dateVal) {
      var parsed = iwv_parseDate_(dateVal);
      dateStr = parsed ? iwv_formatDate_(parsed) : String(dateVal);
    }
    if (!dateSet[dateStr]) continue;

    var visitId = ci.visitId >= 0 ? String(row[ci.visitId]).trim() : '';
    var kind = 'normal';
    if (visitId.indexOf('EV_') === 0) kind = 'event';
    else if (visitId.indexOf('-') >= 0) kind = 'coupled';
    else if (visitId.indexOf('_T_') >= 0) kind = 'trainee';

    result.push({
      visitId: visitId,
      dateStr: dateStr,
      youbi: ci.youbi >= 0 ? iwv_normalizeYoubi_(row[ci.youbi]) : iwv_getYoubiEn_(iwv_parseDate_(dateStr)),
      staffId: ci.staffId >= 0 ? String(row[ci.staffId]).trim() : '',
      staffName: ci.staffName >= 0 ? String(row[ci.staffName]).trim() : '',
      pid: ci.pid >= 0 ? String(row[ci.pid]).trim() : '',
      pname: ci.pname >= 0 ? String(row[ci.pname]).trim() : '',
      area: ci.area >= 0 ? String(row[ci.area]).trim() : '',
      startMin: ci.start >= 0 ? iwv_serialToMinutes_(row[ci.start]) : null,
      endMin: ci.end >= 0 ? iwv_serialToMinutes_(row[ci.end]) : null,
      svcMin: ci.svcMin >= 0 ? (parseInt(row[ci.svcMin]) || 30) : 30,
      timeType: ci.timeType >= 0 ? String(row[ci.timeType]).trim() : '',
      earliestMin: ci.earliest >= 0 ? iwv_serialToMinutes_(row[ci.earliest]) : null,
      latestMin: ci.latest >= 0 ? iwv_serialToMinutes_(row[ci.latest]) : null,
      note: ci.note >= 0 ? String(row[ci.note]).trim() : '',
      kind: kind
    });
  }
  return result;
}

/**
 * 割当不可を読み込み（週でフィルタ）
 */
function iwv_loadUnassigned_(ss, weekDates) {
  var sheet = ss.getSheetByName('割当不可');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var h = data[0];

  var ci = {
    date:      iwv_findHeaderIndex_(h, '日付'),
    youbi:     iwv_findHeaderIndex_(h, '曜日'),
    pid:       iwv_findHeaderIndex_(h, 'patient_id'),
    pname:     iwv_findHeaderIndex_(h, '患者名'),
    needStaff: iwv_findHeaderIndex_(h, '必要スタッフ数'),
    slot:      iwv_findHeaderIndex_(h, '未割当枠'),
    reason:    iwv_findHeaderIndex_(h, '理由')
  };

  var dateSet = {};
  weekDates.forEach(function(d) { dateSet[d] = true; });

  var result = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = ci.date >= 0 ? row[ci.date] : null;
    var dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = iwv_formatDate_(dateVal);
    } else if (dateVal) {
      var parsed = iwv_parseDate_(dateVal);
      dateStr = parsed ? iwv_formatDate_(parsed) : String(dateVal);
    }
    if (!dateSet[dateStr]) continue;

    result.push({
      dateStr: dateStr,
      youbi: ci.youbi >= 0 ? iwv_normalizeYoubi_(row[ci.youbi]) : '',
      pid: ci.pid >= 0 ? String(row[ci.pid]).trim() : '',
      pname: ci.pname >= 0 ? String(row[ci.pname]).trim() : '',
      needStaff: ci.needStaff >= 0 ? (parseInt(row[ci.needStaff]) || 1) : 1,
      slot: ci.slot >= 0 ? (parseInt(row[ci.slot]) || 1) : 1,
      reason: ci.reason >= 0 ? String(row[ci.reason]).trim() : ''
    });
  }
  return result;
}

/**
 * イベントリクエストを読み込み（週でフィルタ、staffId|dateStr でグループ化）
 */
function iwv_loadEvents_(ss, weekDates) {
  var sheet = ss.getSheetByName('イベントリクエスト');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var h = data[0];

  var ci = {
    eventId:  iwv_findHeaderIndex_(h, 'event_id'),
    staffId:  iwv_findHeaderIndex_(h, 'staff_id'),
    date:     iwv_findHeaderIndex_(h, '日付'),
    type:     iwv_findHeaderIndex_(h, 'イベント種別'),
    title:    iwv_findHeaderIndex_(h, 'タイトル'),
    start:    iwv_findHeaderIndex_(h, '開始時刻'),
    end:      iwv_findHeaderIndex_(h, '終了時刻'),
    duration: iwv_findHeaderIndex_(h, '所要時間')
  };

  var dateSet = {};
  weekDates.forEach(function(d) { dateSet[d] = true; });

  var map = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = ci.date >= 0 ? row[ci.date] : null;
    var dateStr = '';
    if (dateVal instanceof Date) dateStr = iwv_formatDate_(dateVal);
    else if (dateVal) {
      var parsed = iwv_parseDate_(dateVal);
      dateStr = parsed ? iwv_formatDate_(parsed) : '';
    }
    if (!dateSet[dateStr]) continue;

    var staffId = ci.staffId >= 0 ? String(row[ci.staffId]).trim() : '';
    var key = staffId + '|' + dateStr;
    if (!map[key]) map[key] = [];
    map[key].push({
      eventId: ci.eventId >= 0 ? String(row[ci.eventId]).trim() : '',
      startMin: ci.start >= 0 ? iwv_serialToMinutes_(row[ci.start]) : null,
      endMin: ci.end >= 0 ? iwv_serialToMinutes_(row[ci.end]) : null,
      duration: ci.duration >= 0 ? (parseInt(row[ci.duration]) || 60) : 60,
      title: ci.title >= 0 ? String(row[ci.title]).trim() : '',
      type: ci.type >= 0 ? String(row[ci.type]).trim() : ''
    });
  }
  return map;
}

/**
 * スタッフ個別変更リクエストを読み込み（週でフィルタ）
 */
function iwv_loadStaffChanges_(ss, weekDates) {
  var sheet = ss.getSheetByName('スタッフ個別変更リクエスト');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var h = data[0];

  var ci = {
    staffId:  iwv_findHeaderIndex_(h, 'staff_id'),
    date:     iwv_findHeaderIndex_(h, '日付'),
    type:     iwv_findHeaderIndex_(h, '制限タイプ'),
    start:    iwv_findHeaderIndex_(h, '開始時刻'),
    end:      iwv_findHeaderIndex_(h, '終了時刻')
  };

  var dateSet = {};
  weekDates.forEach(function(d) { dateSet[d] = true; });

  var map = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = ci.date >= 0 ? row[ci.date] : null;
    var dateStr = '';
    if (dateVal instanceof Date) dateStr = iwv_formatDate_(dateVal);
    else if (dateVal) {
      var parsed = iwv_parseDate_(dateVal);
      dateStr = parsed ? iwv_formatDate_(parsed) : '';
    }
    if (!dateSet[dateStr]) continue;

    var staffId = ci.staffId >= 0 ? String(row[ci.staffId]).trim() : '';
    var key = staffId + '|' + dateStr;
    if (!map[key]) map[key] = [];
    map[key].push({
      restrictionType: ci.type >= 0 ? String(row[ci.type]).trim() : '',
      startTime: ci.start >= 0 ? iwv_serialToMinutes_(row[ci.start]) : null,
      endTime: ci.end >= 0 ? iwv_serialToMinutes_(row[ci.end]) : null
    });
  }
  return map;
}

/**
 * 個別変更リクエストを読み込み（週でフィルタ）
 */
function iwv_loadChangeRequests_(ss, weekDates) {
  var sheet = ss.getSheetByName('個別変更リクエスト');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};
  var h = data[0];

  // 操作カラムは名前が可変のため部分一致検索
  var ci = {
    pid:    iwv_findHeaderIndex_(h, 'patient_id'),
    date:   iwv_findHeaderIndex_(h, '日付'),
    op:     -1,
    start:  iwv_findHeaderIndex_(h, '開始時刻(固定)'),
    end:    iwv_findHeaderIndex_(h, '終了時刻(固定)'),
    timeType: iwv_findHeaderIndex_(h, '時間タイプ'),
    earliest: iwv_findHeaderIndex_(h, '希望最早'),
    latest:   iwv_findHeaderIndex_(h, '希望最遅')
  };
  // 操作カラム: "操作" or "操作種別" or "変更区分" or "操作（キャンセル..."
  for (var i = 0; i < h.length; i++) {
    var hv = String(h[i]).trim();
    if (hv.indexOf('操作') >= 0 || hv.indexOf('変更区分') >= 0) {
      ci.op = i;
      break;
    }
  }
  // 開始時刻(固定)が見つからない場合、"開始時刻"で再検索
  if (ci.start < 0) ci.start = iwv_findHeaderIndex_(h, '開始時刻');
  if (ci.end < 0) ci.end = iwv_findHeaderIndex_(h, '終了時刻');

  var dateSet = {};
  weekDates.forEach(function(d) { dateSet[d] = true; });

  var map = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = ci.date >= 0 ? row[ci.date] : null;
    var dateStr = '';
    if (dateVal instanceof Date) dateStr = iwv_formatDate_(dateVal);
    else if (dateVal) {
      var parsed = iwv_parseDate_(dateVal);
      dateStr = parsed ? iwv_formatDate_(parsed) : '';
    }
    if (!dateSet[dateStr]) continue;

    var pid = ci.pid >= 0 ? String(row[ci.pid]).trim() : '';
    if (!pid) continue;
    var key = pid + '|' + dateStr;
    map[key] = {
      op: ci.op >= 0 ? String(row[ci.op]).trim() : '',
      newStartMin: ci.start >= 0 ? iwv_serialToMinutes_(row[ci.start]) : null,
      newEndMin: ci.end >= 0 ? iwv_serialToMinutes_(row[ci.end]) : null,
      newTimeType: ci.timeType >= 0 ? String(row[ci.timeType]).trim() : '',
      earliestMin: ci.earliest >= 0 ? iwv_serialToMinutes_(row[ci.earliest]) : null,
      latestMin: ci.latest >= 0 ? iwv_serialToMinutes_(row[ci.latest]) : null
    };
  }
  return map;
}

/**
 * 特別訪問週間を読み込み
 */
function iwv_loadSpecialWeek_(ss, weekDates) {
  var result = { mode: '', details: [] };

  // ヘッダシート
  var headerSheet = ss.getSheetByName('特別訪問週間_ヘッダ');
  if (headerSheet) {
    var hData = headerSheet.getDataRange().getValues();
    if (hData.length >= 2) {
      // Mode列を探す
      var hHeaders = hData[0];
      var modeIdx = iwv_findHeaderIndex_(hHeaders, 'Mode');
      if (modeIdx < 0) modeIdx = iwv_findHeaderIndex_(hHeaders, 'モード');
      if (modeIdx >= 0 && hData[1][modeIdx]) {
        result.mode = String(hData[1][modeIdx]).trim().toUpperCase();
      }
    }
  }

  // 明細シート
  var detailSheet = ss.getSheetByName('特別訪問週間_明細');
  if (!detailSheet) return result;
  var data = detailSheet.getDataRange().getValues();
  if (data.length < 2) return result;
  var h = data[0];

  var ci = {
    pid:       iwv_findHeaderIndex_(h, 'patient_id'),
    date:      iwv_findHeaderIndex_(h, '日付'),
    start:     iwv_findHeaderIndex_(h, '開始時刻'),
    end:       iwv_findHeaderIndex_(h, '終了時刻'),
    timeType:  iwv_findHeaderIndex_(h, '時間タイプ'),
    needStaff: iwv_findHeaderIndex_(h, '必要スタッフ数'),
    earliest:  iwv_findHeaderIndex_(h, '希望最早'),
    latest:    iwv_findHeaderIndex_(h, '希望最遅'),
    mode:      iwv_findHeaderIndex_(h, 'Mode')
  };

  var dateSet = {};
  weekDates.forEach(function(d) { dateSet[d] = true; });

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateVal = ci.date >= 0 ? row[ci.date] : null;
    var dateStr = '';
    if (dateVal instanceof Date) dateStr = iwv_formatDate_(dateVal);
    else if (dateVal) {
      var parsed = iwv_parseDate_(dateVal);
      dateStr = parsed ? iwv_formatDate_(parsed) : '';
    }
    if (!dateSet[dateStr]) continue;

    var detailMode = ci.mode >= 0 ? String(row[ci.mode]).trim().toUpperCase() : result.mode;
    result.details.push({
      pid: ci.pid >= 0 ? String(row[ci.pid]).trim() : '',
      dateStr: dateStr,
      startMin: ci.start >= 0 ? iwv_serialToMinutes_(row[ci.start]) : null,
      endMin: ci.end >= 0 ? iwv_serialToMinutes_(row[ci.end]) : null,
      timeType: ci.timeType >= 0 ? String(row[ci.timeType]).trim() : '',
      needStaff: ci.needStaff >= 0 ? (parseInt(row[ci.needStaff]) || 1) : 1,
      earliestMin: ci.earliest >= 0 ? iwv_serialToMinutes_(row[ci.earliest]) : null,
      latestMin: ci.latest >= 0 ? iwv_serialToMinutes_(row[ci.latest]) : null,
      mode: detailMode || result.mode
    });
  }
  return result;
}

/**
 * 分を "HH:mm" 文字列に変換
 */
function iwv_minutesToHHMM_(min) {
  if (min == null || isNaN(min)) return '';
  var h = Math.floor(min / 60);
  var m = min % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

/**
 * 手動変更ログシートにエントリを追記（シート自動作成）
 * @param {Spreadsheet} ss
 * @param {Array[]} logEntries - 2D配列（各行15列）
 */
function iwv_appendChangeLog_(ss, logEntries) {
  var LOG_SHEET_NAME = '手動変更ログ';
  var LOG_HEADERS = [
    '確定日時', 'visit_id', 'patient_id', '患者名', '日付',
    '変更種別', '変更前staff_id', '変更前スタッフ名',
    '変更後staff_id', '変更後スタッフ名',
    '変更前開始', '変更前終了', '変更後開始', '変更後終了', '週開始日'
  ];

  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    logSheet.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold');
  }

  var lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, logEntries.length, logEntries[0].length)
          .setValues(logEntries);
}

// ============================================================
// 確定スケジュール履歴
// ============================================================

/**
 * 割当結果の確定データを「確定スケジュール履歴」シートに保存する。
 * saveSnapshot から自動的に呼び出される。
 * 同じ weekStartStr の既存データは上書き（再確定対応）。
 * @param {string} weekStartStr - "yyyy/MM/dd" 形式の週開始日
 * @returns {Object} { success: boolean, saved: number }
 */
/**
 * 確定スケジュール履歴に保存する。
 * スナップショットコピーから読むことで、割当結果シートが別の週に
 * 上書きされていても正しい週のデータを保存できる。
 * @param {string} weekStartStr - 週開始日
 * @param {string} [snapshotSheetName] - スナップショットシート名（省略時は割当結果から読む）
 */
function iwv_saveConfirmedHistory_(weekStartStr, snapshotSheetName) {
  var ss = iwv_getSpreadsheet_();
  // スナップショットシートがあればそちらから読む（確実）
  var resultSheet = null;
  if (snapshotSheetName) {
    resultSheet = ss.getSheetByName(snapshotSheetName);
  }
  if (!resultSheet) {
    resultSheet = ss.getSheetByName('割当結果');
  }
  if (!resultSheet) throw new Error('割当結果シートが見つかりません');

  // 確定スケジュール履歴シートを取得 or 作成
  var HIST_SHEET_NAME = '確定スケジュール履歴';
  var HIST_HEADERS = [
    '確定日時', '週開始日', 'visit_id', '日付', '曜日',
    'patient_id', '患者名', 'staff_id', 'スタッフ名',
    '開始時刻', '終了時刻', 'サービス時間', '備考'
  ];

  var histSheet = ss.getSheetByName(HIST_SHEET_NAME);
  if (!histSheet) {
    histSheet = ss.insertSheet(HIST_SHEET_NAME);
    histSheet.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);
    histSheet.getRange(1, 1, 1, HIST_HEADERS.length).setFontWeight('bold');
    histSheet.setFrozenRows(1);
  }

  // 割当結果を読み込み
  var data = resultSheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, saved: 0 };

  var headers = data[0];
  var col = {
    visitId:   iwv_findHeaderIndex_(headers, 'visit_id'),
    date:      iwv_findHeaderIndex_(headers, '日付'),
    youbi:     iwv_findHeaderIndex_(headers, '曜日'),
    staffId:   iwv_findHeaderIndex_(headers, 'staff_id'),
    staffName: iwv_findHeaderIndex_(headers, 'スタッフ名'),
    pid:       iwv_findHeaderIndex_(headers, 'patient_id'),
    pname:     iwv_findHeaderIndex_(headers, '患者名'),
    start:     iwv_findHeaderIndex_(headers, '開始時刻'),
    end:       iwv_findHeaderIndex_(headers, '終了時刻'),
    svcMin:    iwv_findHeaderIndex_(headers, 'サービス時間'),
    note:      iwv_findHeaderIndex_(headers, '備考')
  };

  var now = new Date();
  var rows = [];

  // 週の日付範囲を算出（weekStartStr から7日間）
  var tz = ss.getSpreadsheetTimeZone();
  var weekStartDate = iwv_parseDate_(weekStartStr);
  var weekDateSet = {};
  if (weekStartDate) {
    for (var wi = 0; wi < 7; wi++) {
      var wd = new Date(weekStartDate.getTime() + wi * 86400000);
      weekDateSet[Utilities.formatDate(wd, tz, 'yyyy/MM/dd')] = true;
    }
  }

  var skippedOutOfRange = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    // visit_id が空 or イベント行 (EV_) はスキップ
    var vid = col.visitId >= 0 ? String(r[col.visitId] || '') : '';
    if (!vid || vid.indexOf('EV_') === 0) continue;
    // 未割当（staff_id が空）はスキップ
    var sid = col.staffId >= 0 ? String(r[col.staffId] || '') : '';
    if (!sid) continue;
    // 日付が対象週の範囲内か確認
    if (weekStartDate && col.date >= 0) {
      var rowDate = r[col.date];
      var rowDateStr = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, tz, 'yyyy/MM/dd') : String(rowDate || '');
      if (rowDateStr && !weekDateSet[rowDateStr]) {
        skippedOutOfRange++;
        continue; // この週に含まれない日付はスキップ
      }
    }

    rows.push([
      now,
      weekStartStr,
      vid,
      col.date >= 0 ? r[col.date] : '',
      col.youbi >= 0 ? String(r[col.youbi] || '') : '',
      col.pid >= 0 ? String(r[col.pid] || '') : '',
      col.pname >= 0 ? String(r[col.pname] || '') : '',
      sid,
      col.staffName >= 0 ? String(r[col.staffName] || '') : '',
      col.start >= 0 ? r[col.start] : '',
      col.end >= 0 ? r[col.end] : '',
      col.svcMin >= 0 ? r[col.svcMin] : '',
      col.note >= 0 ? String(r[col.note] || '') : ''
    ]);
  }

  if (skippedOutOfRange > 0) {
    Logger.log('確定スケジュール履歴: ' + skippedOutOfRange + '件を週範囲外としてスキップ');
  }

  if (rows.length > 0) {
    // 同じ weekStartStr の既存行を削除（再確定時の上書き）
    var existingData = histSheet.getDataRange().getValues();
    var rowsToDelete = [];
    for (var j = existingData.length - 1; j >= 1; j--) {
      if (String(existingData[j][1]) === weekStartStr) {
        rowsToDelete.push(j + 1); // 1-based row number
      }
    }
    // 逆順で削除してインデックスのずれを防ぐ
    for (var d = 0; d < rowsToDelete.length; d++) {
      histSheet.deleteRow(rowsToDelete[d]);
    }

    // 新しい行を追記
    histSheet.getRange(histSheet.getLastRow() + 1, 1, rows.length, HIST_HEADERS.length)
             .setValues(rows);
  }

  return { success: true, saved: rows.length };
}

/**
 * 患者のローテーション履歴を取得する（確定スケジュール履歴から）。
 * 過去N週分のスタッフ割当情報を返す。
 * @param {string} patientId - 患者ID
 * @param {number} [weeksBack=4] - 遡る週数（デフォルト4）
 * @returns {Object} { history: Array<{weekStart, date, staffId, staffName}> }
 */
function getPatientRotationHistory(patientId, weeksBack) {
  var ss = iwv_getSpreadsheet_();
  var sheet = ss.getSheetByName('確定スケジュール履歴');
  if (!sheet) return { history: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { history: [] };

  weeksBack = weeksBack || 4;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - (weeksBack * 7));

  var headers = data[0];
  var colPid = iwv_findHeaderIndex_(headers, 'patient_id');
  var colConfirm = iwv_findHeaderIndex_(headers, '確定日時');
  var colWeekStart = iwv_findHeaderIndex_(headers, '週開始日');
  var colDate = iwv_findHeaderIndex_(headers, '日付');
  var colStaffId = iwv_findHeaderIndex_(headers, 'staff_id');
  var colStaffName = iwv_findHeaderIndex_(headers, 'スタッフ名');

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var pid = colPid >= 0 ? String(row[colPid] || '') : '';
    if (pid !== patientId) continue;

    var confirmDate = colConfirm >= 0 ? row[colConfirm] : null;
    if (confirmDate instanceof Date && confirmDate < cutoff) continue;

    results.push({
      weekStart: colWeekStart >= 0 ? String(row[colWeekStart] || '') : '',
      date: colDate >= 0 ? row[colDate] : '',
      staffId: colStaffId >= 0 ? String(row[colStaffId] || '') : '',
      staffName: colStaffName >= 0 ? String(row[colStaffName] || '') : ''
    });
  }

  return { history: results };
}
