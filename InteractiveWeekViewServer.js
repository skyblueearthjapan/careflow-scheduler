/**
 * InteractiveWeekViewServer.js
 * GAS サーバー側ロジック: インタラクティブ週ビュー
 *
 * Public functions (callable via google.script.run):
 *   - openInteractiveWeekView(weekStartStr)
 *   - getInteractiveWeekData(weekStartStr)
 *   - commitChanges(weekStartStr, changesJson)
 *   - saveSnapshot(weekStartStr)
 */

// ============================================================
// Public Functions
// ============================================================

/**
 * インタラクティブ週ビューダイアログを開く
 * @param {string} weekStartStr - "yyyy/MM/dd" 形式の週開始日
 */
function openInteractiveWeekView(weekStartStr) {
  var template = HtmlService.createTemplateFromFile('InteractiveWeekView');
  template.weekStartStr = weekStartStr || '';
  var html = template.evaluate()
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '週ビュー インタラクティブ編集');
}

/**
 * インタラクティブ週ビューに必要な全データを取得
 * @param {string} weekStartStr - "yyyy/MM/dd" 形式の週開始日
 * @returns {Object} 全データを含むオブジェクト
 */
function getInteractiveWeekData(weekStartStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var weekStart = iwv_parseDate_(weekStartStr);
  if (!weekStart) throw new Error('無効な週開始日: ' + weekStartStr);

  var weekDates = [];
  for (var d = 0; d < 7; d++) {
    var dt = new Date(weekStart);
    dt.setDate(dt.getDate() + d);
    weekDates.push(iwv_formatDate_(dt));
  }

  return {
    weekStartStr: weekStartStr,
    weekDates: weekDates,
    staffList: iwv_loadStaffMaster_(ss),
    patientMap: iwv_loadPatientMaster_(ss),
    assignments: iwv_loadAssignments_(ss, weekDates),
    unassigned: iwv_loadUnassigned_(ss, weekDates),
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

  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  var applied = 0;
  changes.forEach(function(change) {
    var found = false;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][col.visitId]) === String(change.visitId)) {
        found = true;
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
        // 備考に手動変更マーク追加
        var note = String(data[r][col.note] || '');
        if (note.indexOf('[手動変更]') === -1) {
          if (note && !note.endsWith(' ')) note += ' ';
          note += '[手動変更]';
          data[r][col.note] = note;
        }
        applied++;
        break;
      }
    }
    if (!found) {
      Logger.log('commitChanges: visit_id "' + change.visitId + '" が割当結果に見つかりません');
    }
  });

  // シートに書き戻し
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  return {
    success: true,
    resultSheetName: resultName,
    weekSheetName: weekName,
    label: label
  };
}

// ============================================================
// Private Helper Functions (iwv_ prefix)
// ============================================================

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
    insurance:  iwv_findHeaderIndex_(h, '保険区分')
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
      insuranceType: ci.insurance >= 0 ? String(row[ci.insurance]).trim() : ''
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
