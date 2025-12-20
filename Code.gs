/**
 * 訪問看護 自動スケジューリング - Webアプリ
 * 第1段階：週ビュー表示 + GAS実行ボタン群
 */

// スクリプトプロパティから設定を取得
const SS_ID = PropertiesService.getScriptProperties().getProperty('SS_ID');
const WEEKVIEW_SHEET = PropertiesService.getScriptProperties().getProperty('SHEET_WEEKVIEW') || '週ビュー';
const LOG_SHEET = PropertiesService.getScriptProperties().getProperty('SHEET_LOG') || '実行ログ';
const INPUT_APP_URL = PropertiesService.getScriptProperties().getProperty('INPUT_APP_URL') || '';

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
 */
function getWeekView() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(WEEKVIEW_SHEET);

    if (!sheet) {
      throw new Error(`シート「${WEEKVIEW_SHEET}」が見つかりません`);
    }

    const lastRow = findLastStaffRow_(sheet);
    const lastCol = 8;

    const values = sheet.getRange(1, 1, Math.max(1, lastRow), lastCol).getValues();
    const header = values[0];
    const rows = values.slice(1).map(r => ({
      staff: r[0],
      cells: r.slice(1)
    })).filter(x => String(x.staff || '').trim() !== '');

    return {
      success: true,
      header,
      rows,
      meta: { lastRow, timestamp: new Date().toISOString() }
    };
  } catch (e) {
    console.error('getWeekView error:', e);
    return { success: false, error: e.message };
  }
}

function findLastStaffRow_(sheet) {
  const max = sheet.getMaxRows();
  if (max <= 1) return 1;
  const colA = sheet.getRange(2, 1, max - 1, 1).getValues();
  let last = 1;
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || '').trim() !== '') last = i + 2;
  }
  return last;
}

// ============================================================
// 汎用シートテーブル取得API（第1.5段階）
// ============================================================

/**
 * 汎用シートデータ取得
 * @param {string} sheetName - シート名
 * @param {Object} opt - オプション
 * @param {number} opt.limit - 行数上限（デフォルト300）
 * @param {boolean} opt.filterThisWeek - 今週フィルタを適用するか
 * @param {string} opt.dateColName - 日付列名（デフォルト'日付'）
 */
function getSheetTable(sheetName, opt) {
  opt = opt || {};
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('シートが見つかりません: ' + sheetName);

  const values = sh.getDataRange().getValues();
  if (values.length === 0) return { headers: [], rows: [] };

  const headers = values[0];
  let rows = values.slice(1);

  // 空行を除去
  rows = rows.filter(r => r.some(cell => cell !== '' && cell !== null && cell !== undefined));

  // 今週フィルタ（"日付"列がある場合だけ適用）
  if (opt.filterThisWeek) {
    const tz = ss.getSpreadsheetTimeZone();
    const dateColName = opt.dateColName || '日付';
    const idxDate = headers.indexOf(dateColName);

    if (idxDate >= 0) {
      const { start, end } = getThisWeekRange_(tz);
      rows = rows.filter(r => {
        const d = r[idxDate];
        return (d instanceof Date) && d >= start && d <= end;
      });
    }
  }

  // 行数制限（末尾から）
  const limit = opt.limit || 300;
  if (rows.length > limit) rows = rows.slice(rows.length - limit);

  // Date型をISO文字列に変換（JSON転送用）
  rows = rows.map(row => row.map(cell => {
    if (cell instanceof Date) {
      return { _type: 'date', value: cell.toISOString() };
    }
    return cell;
  }));

  return { headers, rows, rowCount: rows.length };
}

/**
 * 今週の範囲を取得（月曜開始）
 */
function getThisWeekRange_(tz) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const day = today.getDay();
  const diffToMonday = (day + 6) % 7;
  const start = new Date(today);
  start.setDate(today.getDate() - diffToMonday);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  end.setHours(23, 59, 59, 999);
  return { start, end };
}

// ============================================================
// GAS実行ラッパー関数（UIから呼ばれる）
// ============================================================

function runGenerateWeeklyRequests() {
  const startTime = new Date();
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const result = 週間リクエストを生成_(ss);
    const message = result.message || '週間リクエストの生成が完了しました';
    logExecution_('週間リクエスト生成', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    logExecution_('週間リクエスト生成', false, e.message, startTime);
    return { success: false, error: e.message };
  }
}

function runCreateAssignments() {
  const startTime = new Date();
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const result = 割当結果を作成_(ss);
    const message = result.message || '割当結果の作成が完了しました';
    logExecution_('割当結果作成', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    logExecution_('割当結果作成', false, e.message, startTime);
    return { success: false, error: e.message };
  }
}

function runUpdateWeekView() {
  const startTime = new Date();
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    週ビューを更新_(ss);
    const message = '週ビューの更新が完了しました';
    logExecution_('週ビュー更新', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    logExecution_('週ビュー更新', false, e.message, startTime);
    return { success: false, error: e.message };
  }
}

function runCreateRouteSummary() {
  const startTime = new Date();
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const result = ルートサマリを作成_(ss);
    const message = result.message || 'ルートサマリの作成が完了しました';
    logExecution_('ルートサマリ作成', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    logExecution_('ルートサマリ作成', false, e.message, startTime);
    return { success: false, error: e.message };
  }
}

function runUpdateLocation() {
  const startTime = new Date();
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    位置情報を更新_(ss);
    const message = '位置情報の更新が完了しました';
    logExecution_('位置情報更新', true, message, startTime);
    return { success: true, message };
  } catch (e) {
    logExecution_('位置情報更新', false, e.message, startTime);
    return { success: false, error: e.message };
  }
}

// ============================================================
// ログ機能
// ============================================================

function logExecution_(actionName, success, message, startTime) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let logSheet = ss.getSheetByName(LOG_SHEET);

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

function getExecutionLogs(limit = 20) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const logSheet = ss.getSheetByName(LOG_SHEET);

    if (!logSheet) return { success: true, logs: [] };

    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) return { success: true, logs: [] };

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
    return { success: false, error: e.message };
  }
}

function checkConfiguration() {
  const issues = [];
  if (!SS_ID) issues.push('SS_ID（スプレッドシートID）が設定されていません');

  if (SS_ID) {
    try {
      const ss = SpreadsheetApp.openById(SS_ID);
      const sheet = ss.getSheetByName(WEEKVIEW_SHEET);
      if (!sheet) issues.push(`シート「${WEEKVIEW_SHEET}」が見つかりません`);
    } catch (e) {
      issues.push(`スプレッドシートにアクセスできません: ${e.message}`);
    }
  }

  return {
    success: issues.length === 0,
    issues,
    config: { SS_ID: SS_ID ? '設定済み' : '未設定', WEEKVIEW_SHEET, LOG_SHEET }
  };
}

/**
 * アプリリンク情報を取得
 */
function getAppLinks() {
  return {
    inputAppUrl: INPUT_APP_URL || null,
    hasInputApp: !!INPUT_APP_URL
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

function buildSlotPrefs(specIdStr, specTypeStr, needStaff) {
  var ids = parseIdList(specIdStr);
  var types = parseTypeList(specTypeStr);

  if (!ids.length && !types.length) {
    var result = [];
    for (var i = 0; i < needStaff; i++) {
      result.push({ type: '', ids: {}, poolMode: false });
    }
    return result;
  }

  if (types.length >= needStaff) {
    var result = [];
    for (var i = 0; i < needStaff; i++) {
      var idSet = {};
      if (ids[i]) {
        idSet[ids[i]] = true;
      } else {
        ids.forEach(function(id) { idSet[id] = true; });
      }
      result.push({ type: types[i] || '', ids: idSet, poolMode: false });
    }
    return result;
  }

  if (types.length === 1) {
    var t = types[0];
    var idSet = {};
    ids.forEach(function(id) { idSet[id] = true; });
    var result = [];
    for (var i = 0; i < needStaff; i++) {
      result.push({ type: t, ids: idSet, poolMode: true });
    }
    return result;
  }

  if (!types.length && ids.length) {
    var idSet = {};
    ids.forEach(function(id) { idSet[id] = true; });
    var result = [];
    for (var i = 0; i < needStaff; i++) {
      result.push({ type: '優先', ids: idSet, poolMode: true });
    }
    return result;
  }

  var result = [];
  for (var i = 0; i < needStaff; i++) {
    result.push({ type: '', ids: {}, poolMode: false });
  }
  return result;
}

function applyStaffPreferenceToCandidates(candidates, slotPref, ngSet) {
  var arr = candidates.filter(function(c) { return !ngSet[c.staff.id]; });
  var t = (slotPref.type || '').trim();
  var ids = slotPref.ids || {};
  var hasIds = false;
  for (var k in ids) { hasIds = true; break; }
  if (!hasIds || !t) return arr;
  if (t === '必須') return arr.filter(function(c) { return ids[c.staff.id]; });
  if (t === '優先') {
    var preferred = arr.filter(function(c) { return ids[c.staff.id]; });
    var others = arr.filter(function(c) { return !ids[c.staff.id]; });
    return preferred.concat(others);
  }
  return arr;
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
      pData.forEach(r => {
        const id = r[pIdxId];
        if (!id) return;
        patientGenderMap[id] = pIdxGen >= 0 ? (r[pIdxGen] || '') : '';
      });
    }
  }

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

  staffList.forEach((st, rIndex) => {
    for (let i = 0; i < 7; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      const targetDateStr = Utilities.formatDate(d, tz, 'yyyy/MM/dd');

      const visits = weekData.filter(row => {
        const d2 = row[idxDate];
        const ds2 = Utilities.formatDate(d2, tz, 'yyyy/MM/dd');
        const sid   = row[idxStaffId] || '';
        const sname = row[idxStaff]   || '';
        const key   = sid || sname;
        return key === (st.id || st.name) && ds2 === targetDateStr;
      });

      visits.sort((a,b) => a[idxStart] - b[idxStart]);

      const lines = visits.map(v => {
        const startVal = v[idxStart];
        const endVal   = v[idxEnd];
        const pid      = v[idxPid] || '';
        const pname    = v[idxPatient] || '';
        const pGender  = pid ? (patientGenderMap[pid] || '') : '';
        const vid      = (idxVisitId >= 0) ? (v[idxVisitId] || '') : '';
        const noteVal  = (idxNote >= 0) ? (v[idxNote] || '') : '';

        const isTwo = (String(vid).indexOf('-') >= 0) || String(noteVal).indexOf('同時訪問') >= 0;
        const mark = isTwo ? '👥 ' : '';

        function formatTime(val) {
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

        const stime = formatTime(startVal);
        const etime = formatTime(endVal);

        let pidPart = '';
        if (pid) {
          pidPart = pid;
          if (pGender) pidPart += '（' + pGender + '）';
          pidPart += ' ';
        }

        if (!stime && !etime) return mark + pidPart + pname;
        return mark + stime + '〜' + etime + ' ' + pidPart + pname;
      });

      const cellText = lines.join('\n');
      if (cellText) {
        const cell = viewSheet.getRange(2 + rIndex, 2 + i);
        cell.setValue(cellText);
        cell.setWrap(true);
      }
    }
  });
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

  function toMinutes(v) {
    if (typeof v === 'number') return Math.round(v * 24 * 60);
    else if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
    else return null;
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

  var staffDateMap = {};
  resultRows.forEach(function(r, i){
    var d = r[1], staffId = r[3], pid = r[5];
    if (!staffId || !(d instanceof Date) || !pid) return;
    var key = staffId + '|' + Utilities.formatDate(d, tz, 'yyyy/MM/dd');
    if (!staffDateMap[key]) staffDateMap[key] = [];
    staffDateMap[key].push(i);
  });

  Object.keys(staffDateMap).forEach(function(key){
    var indexList = staffDateMap[key];
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

  // 時刻自動調整
  Object.keys(staffDateMap).forEach(function(key){
    var idxList = staffDateMap[key];
    idxList.sort(function(aIdx, bIdx){ return toMinutes(resultRows[aIdx][8]) - toMinutes(resultRows[bIdx][8]); });
    var currentEndMin = null;
    idxList.forEach(function(rIdx){
      var row = resultRows[rIdx];
      var timeType = row[11];
      var svcMin = Number(row[10]) || 0;
      if (!svcMin) return;
      var earliestMin = row[12] ? toMinutes(row[12]) : null;
      var latestMin = row[13] ? toMinutes(row[13]) : null;
      var moveMin = Number(moveMinArr[rIdx]) || 0;
      var gapMin = moveMin + EXTRA_BUFFER_MIN;

      if (timeType === '固定') {
        var fixedStartMin = toMinutes(row[8]);
        if (fixedStartMin == null) {
          if (currentEndMin != null) fixedStartMin = currentEndMin + gapMin;
          else if (earliestMin != null) fixedStartMin = earliestMin;
        }
        var fixedEndMin = fixedStartMin + svcMin;
        row[8] = fixedStartMin / (24 * 60);
        row[9] = fixedEndMin / (24 * 60);
        currentEndMin = fixedEndMin;
        return;
      }

      var baseStartMin = toMinutes(row[8]);
      var startCandidate = baseStartMin != null ? baseStartMin : (earliestMin != null ? earliestMin : currentEndMin);
      if (currentEndMin != null) startCandidate = Math.max(startCandidate || 0, currentEndMin + gapMin);
      if (earliestMin != null) startCandidate = Math.max(startCandidate, earliestMin);
      var startMin = startCandidate;
      var endMin = startMin + svcMin;
      if (latestMin != null && endMin > latestMin) row[14] = (row[14] || '') + ' / 希望時間帯内に収まらない可能性あり';
      row[8] = startMin / (24 * 60);
      row[9] = endMin / (24 * 60);
      currentEndMin = endMin;
    });
  });

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
