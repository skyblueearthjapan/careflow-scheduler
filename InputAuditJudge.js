/**
 * 入力データ監査 - 判定ロジック
 *
 * スプレッドシートの各入力シートを行ごと・セルごとにチェックし、
 * NG/WARNの判定結果を返す。
 */

// ============================================================
// メイン判定関数
// ============================================================

/**
 * 指定シートの全データを監査し、セルごとの判定結果を返す
 * @param {string} sheetName - シート名
 * @return {Object} { summary, cellResults, rowResults }
 */
function inputAudit_judgeSheet(sheetName) {
  var data = inputAudit_getSheetData_(sheetName);
  if (!data) {
    return { summary: { total: 0, ng: 0, warn: 0, ok: 0 }, cellResults: [], rowResults: [] };
  }

  var headers = data.headers;
  var rows = data.rows;
  var cellResults = [];
  var rowResults = [];
  var totalNg = 0;
  var totalWarn = 0;
  var totalOk = 0;

  for (var rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    var row = rows[rowIdx];

    // 空行スキップ
    if (inputAudit_isEmptyRow_(row)) continue;

    var issues = inputAudit_judgeRow_(sheetName, headers, row, rowIdx);

    var rowNg = 0;
    var rowWarn = 0;
    for (var i = 0; i < issues.length; i++) {
      cellResults.push(issues[i]);
      if (issues[i].level === 'NG') rowNg++;
      else if (issues[i].level === 'WARN') rowWarn++;
    }

    var rowLevel = 'OK';
    if (rowNg > 0) { rowLevel = 'NG'; totalNg++; }
    else if (rowWarn > 0) { rowLevel = 'WARN'; totalWarn++; }
    else { totalOk++; }

    rowResults.push({ rowIdx: rowIdx, level: rowLevel, issues: issues });
  }

  return {
    summary: { total: rows.length, ng: totalNg, warn: totalWarn, ok: totalOk },
    cellResults: cellResults,
    rowResults: rowResults
  };
}

// ============================================================
// 行判定ディスパッチ
// ============================================================

function inputAudit_judgeRow_(sheetName, headers, row, rowIdx) {
  switch (sheetName) {
    case '患者マスタ':
      return inputAudit_judgePatientRow_(headers, row, rowIdx);
    case 'スタッフマスタ':
      return inputAudit_judgeStaffRow_(headers, row, rowIdx);
    case '個別変更リクエスト':
      return inputAudit_judgeChangeRow_(headers, row, rowIdx);
    case 'スタッフ個別変更リクエスト':
      return inputAudit_judgeStaffChangeRow_(headers, row, rowIdx);
    case 'イベントリクエスト':
      return inputAudit_judgeEventRow_(headers, row, rowIdx);
    case 'スタッフ同行割付':
      return inputAudit_judgeMentorRow_(headers, row, rowIdx);
    case '週間訪問パターン':
      return inputAudit_judgePatternRow_(headers, row, rowIdx);
    default:
      return [];
  }
}

// ============================================================
// 患者マスタ判定
// ============================================================

function inputAudit_judgePatientRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var pid = inputAudit_getVal_(row, h['patient_id']);
  var name = inputAudit_getVal_(row, h['患者名']);
  var status = inputAudit_getVal_(row, h['稼働状況']);
  var weeklyCount = inputAudit_getVal_(row, h['週訪問回数']);
  var prefDays = inputAudit_getVal_(row, h['希望曜日']);
  var svcTime = inputAudit_getVal_(row, h['サービス時間']);
  var timeType = inputAudit_getVal_(row, h['時間タイプ']);
  var startTime = inputAudit_getVal_(row, h['希望時間帯（開始）']);
  var endTime = inputAudit_getVal_(row, h['希望時間帯（終了）']);
  var lat = inputAudit_getVal_(row, h['緯度']);
  var lng = inputAudit_getVal_(row, h['経度']);
  var sexLimit = inputAudit_getVal_(row, h['性別制限']);
  var needStaff = inputAudit_getVal_(row, h['必要スタッフ数']);
  var specStaff = inputAudit_getVal_(row, h['指定スタッフID']);
  var specType = inputAudit_getVal_(row, h['指定タイプ']);
  var contPref = inputAudit_getVal_(row, h['継続希望']);

  // P-E01
  if (!pid) issues.push(inputAudit_issue_(rowIdx, h['patient_id'], 'patient_id', 'NG', 'P-E01', '患者IDが未設定です'));
  // P-E02
  if (!name) issues.push(inputAudit_issue_(rowIdx, h['患者名'], '患者名', 'NG', 'P-E02', '患者名が未入力です'));
  // P-E03
  if (!status) issues.push(inputAudit_issue_(rowIdx, h['稼働状況'], '稼働状況', 'NG', 'P-E03', '稼働状況が未設定です。算出対象か判定できません'));
  // P-E11: 稼働状況の値が不正
  if (status && !inputAudit_isValidChoice(status, '稼働状況')) {
    issues.push(inputAudit_issue_(rowIdx, h['稼働状況'], '稼働状況', 'NG', 'P-E11', '稼働状況の値が不正です（「' + status + '」）'));
  }

  // 非稼働患者は以降のチェックをスキップ（稼働以外は算出対象外）
  if (status && status !== '稼働') {
    return issues;
  }

  // P-E04
  var weeklyNum = audit_toNumber_(weeklyCount, 0);
  if (!weeklyCount || weeklyNum <= 0) issues.push(inputAudit_issue_(rowIdx, h['週訪問回数'], '週訪問回数', 'NG', 'P-E04', '週訪問回数が未設定です。訪問リクエストを生成できません'));
  // P-E05
  var parsedDays = audit_parseDays_(prefDays);
  if (parsedDays.length === 0) issues.push(inputAudit_issue_(rowIdx, h['希望曜日'], '希望曜日', 'NG', 'P-E05', '希望曜日が未選択です。訪問日を決定できません'));
  // P-E06
  if (!svcTime) issues.push(inputAudit_issue_(rowIdx, h['サービス時間'], 'サービス時間', 'NG', 'P-E06', 'サービス時間が未入力です。時間枠を計算できません'));
  // P-E07: 固定なのに開始/終了が空
  if (timeType === '固定' && !startTime) {
    issues.push(inputAudit_issue_(rowIdx, h['希望時間帯（開始）'], '希望時間帯（開始）', 'NG', 'P-E07', '時間タイプが「固定」ですが開始時刻が未入力です'));
  }
  if (timeType === '固定' && !endTime) {
    issues.push(inputAudit_issue_(rowIdx, h['希望時間帯（終了）'], '希望時間帯（終了）', 'NG', 'P-E07', '時間タイプが「固定」ですが終了時刻が未入力です'));
  }
  // P-E08: 時間帯なのに開始/終了が空
  if (timeType === '時間帯' && !startTime) {
    issues.push(inputAudit_issue_(rowIdx, h['希望時間帯（開始）'], '希望時間帯（開始）', 'NG', 'P-E08', '時間タイプが「時間帯」ですが開始時刻が未入力です'));
  }
  if (timeType === '時間帯' && !endTime) {
    issues.push(inputAudit_issue_(rowIdx, h['希望時間帯（終了）'], '希望時間帯（終了）', 'NG', 'P-E08', '時間タイプが「時間帯」ですが終了時刻が未入力です'));
  }
  // P-E09: 開始 >= 終了
  if (startTime && endTime) {
    var startMin = audit_parseTimeToMin_(startTime);
    var endMin = audit_parseTimeToMin_(endTime);
    if (startMin !== null && endMin !== null && startMin >= endMin) {
      issues.push(inputAudit_issue_(rowIdx, h['希望時間帯（開始）'], '希望時間帯（開始）', 'NG', 'P-E09', '開始時刻が終了時刻以降になっています'));
    }
  }
  // P-E10: 曜日数 < 週訪問回数
  if (parsedDays.length > 0 && weeklyNum > 0 && parsedDays.length < weeklyNum) {
    issues.push(inputAudit_issue_(rowIdx, h['希望曜日'], '希望曜日', 'NG', 'P-E10', '希望曜日数（' + parsedDays.length + '日）が週訪問回数（' + weeklyNum + '回）より少ないです'));
  }
  // P-E12: 時間タイプの値が不正
  if (timeType && !inputAudit_isValidChoice(timeType, '時間タイプ')) {
    issues.push(inputAudit_issue_(rowIdx, h['時間タイプ'], '時間タイプ', 'NG', 'P-E12', '時間タイプの値が不正です（「' + timeType + '」）'));
  }
  // P-W01: 位置情報なし
  if (!lat || !lng) issues.push(inputAudit_issue_(rowIdx, h['緯度'], '緯度', 'WARN', 'P-W01', '住所（位置情報）が未設定です。ルート最適化・距離計算に影響します'));
  // P-W02: 時間タイプ空
  if (!timeType) issues.push(inputAudit_issue_(rowIdx, h['時間タイプ'], '時間タイプ', 'WARN', 'P-W02', '時間タイプが未設定です。終日扱いとなり精度が低下します'));
  // P-W03: 継続希望空
  if (!contPref) issues.push(inputAudit_issue_(rowIdx, h['継続希望'], '継続希望', 'WARN', 'P-W03', '継続希望が未設定です。「どちらでも」として扱われます'));
  // P-W04: 指定スタッフありなのに指定タイプ空
  if (specStaff && !specType) {
    issues.push(inputAudit_issue_(rowIdx, h['指定タイプ'], '指定タイプ', 'WARN', 'P-W04', '指定スタッフIDが設定されていますが指定タイプが未選択です'));
  }
  // P-W05: 必要スタッフ数空
  if (!needStaff) issues.push(inputAudit_issue_(rowIdx, h['必要スタッフ数'], '必要スタッフ数', 'WARN', 'P-W05', '必要スタッフ数が未設定です。デフォルト（1名）で処理されます'));
  // P-W06: 性別制限不正
  if (sexLimit && !inputAudit_isValidChoice(sexLimit, '性別制限')) {
    issues.push(inputAudit_issue_(rowIdx, h['性別制限'], '性別制限', 'WARN', 'P-W06', '性別制限の値が不正です（「' + sexLimit + '」）'));
  }
  // P-W07: 継続希望不正
  if (contPref && !inputAudit_isValidChoice(contPref, '継続希望')) {
    issues.push(inputAudit_issue_(rowIdx, h['継続希望'], '継続希望', 'WARN', 'P-W07', '継続希望の値が不正です（「' + contPref + '」）'));
  }

  return issues;
}

// ============================================================
// スタッフマスタ判定
// ============================================================

function inputAudit_judgeStaffRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var sid = inputAudit_getVal_(row, h['staff_id']);
  var name = inputAudit_getVal_(row, h['スタッフ名']);
  var gender = inputAudit_getVal_(row, h['性別']);
  var shiftStart = inputAudit_getVal_(row, h['シフト開始']);
  var shiftEnd = inputAudit_getVal_(row, h['シフト終了']);
  var workDays = inputAudit_getVal_(row, h['勤務曜日']);
  var maxPerDay = inputAudit_getVal_(row, h['最大訪問件数/日']);
  var lat = inputAudit_getVal_(row, h['緯度']);
  var allocPref = inputAudit_getVal_(row, h['割当量']);

  // S-E01
  if (!sid) issues.push(inputAudit_issue_(rowIdx, h['staff_id'], 'staff_id', 'NG', 'S-E01', 'スタッフIDが未設定です'));
  // S-E02
  if (!name) issues.push(inputAudit_issue_(rowIdx, h['スタッフ名'], 'スタッフ名', 'NG', 'S-E02', 'スタッフ名が未入力です'));
  // S-E03
  var parsedWork = audit_parseDays_(workDays);
  if (parsedWork.length === 0) issues.push(inputAudit_issue_(rowIdx, h['勤務曜日'], '勤務曜日', 'NG', 'S-E03', '勤務曜日が未選択です。割当先として使用できません'));
  // S-E04
  if (!shiftStart) issues.push(inputAudit_issue_(rowIdx, h['シフト開始'], 'シフト開始', 'NG', 'S-E04', 'シフト開始時刻が未設定です'));
  // S-E05
  if (!shiftEnd) issues.push(inputAudit_issue_(rowIdx, h['シフト終了'], 'シフト終了', 'NG', 'S-E05', 'シフト終了時刻が未設定です'));
  // S-E06: 開始 >= 終了
  if (shiftStart && shiftEnd) {
    var sMin = audit_parseTimeToMin_(shiftStart);
    var eMin = audit_parseTimeToMin_(shiftEnd);
    if (sMin !== null && eMin !== null && sMin >= eMin) {
      issues.push(inputAudit_issue_(rowIdx, h['シフト開始'], 'シフト開始', 'NG', 'S-E06', 'シフト開始が終了以降になっています'));
    }
  }
  // S-W01
  if (!gender) issues.push(inputAudit_issue_(rowIdx, h['性別'], '性別', 'WARN', 'S-W01', '性別が未設定です。性別制限のある患者に割当できません'));
  // S-W05: 性別不正
  if (gender && !inputAudit_isValidChoice(gender, '性別')) {
    issues.push(inputAudit_issue_(rowIdx, h['性別'], '性別', 'WARN', 'S-W05', '性別の値が不正です（「' + gender + '」）'));
  }
  // S-W02
  if (!lat) issues.push(inputAudit_issue_(rowIdx, h['緯度'], '緯度', 'WARN', 'S-W02', '拠点住所（位置情報）が未設定です。距離計算に影響します'));
  // S-W03
  if (!maxPerDay) issues.push(inputAudit_issue_(rowIdx, h['最大訪問件数/日'], '最大訪問件数/日', 'WARN', 'S-W03', '最大訪問件数が未設定です。デフォルト（999=無制限）で処理されます'));
  // S-W04
  if (allocPref && !inputAudit_isValidChoice(allocPref, '割当量')) {
    issues.push(inputAudit_issue_(rowIdx, h['割当量'], '割当量', 'WARN', 'S-W04', '割当量の値が不正です（「' + allocPref + '」）'));
  }

  return issues;
}

// ============================================================
// 個別変更リクエスト判定
// ============================================================

function inputAudit_judgeChangeRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var pid = inputAudit_getVal_(row, h['patient_id']);
  var date = inputAudit_getVal_(row, h['日付']);
  var op = inputAudit_getVal_(row, h['操作']);
  var timeType = inputAudit_getVal_(row, h['時間タイプ']);
  var startFixed = inputAudit_getVal_(row, h['開始時刻(固定)']);
  var endFixed = inputAudit_getVal_(row, h['終了時刻(固定)']);
  var earliest = inputAudit_getVal_(row, h['希望最早']);
  var latest = inputAudit_getVal_(row, h['希望最遅']);

  // C-E01
  if (!pid) issues.push(inputAudit_issue_(rowIdx, h['patient_id'], 'patient_id', 'NG', 'C-E01', '患者IDが未設定です'));
  // C-E02
  if (!date) issues.push(inputAudit_issue_(rowIdx, h['日付'], '日付', 'NG', 'C-E02', '日付が未入力です'));
  // C-E03
  if (!op) issues.push(inputAudit_issue_(rowIdx, h['操作'], '操作', 'NG', 'C-E03', '操作が未選択です'));
  // C-E06: 操作の値が不正
  if (op && !inputAudit_isValidChoice(op, '操作')) {
    issues.push(inputAudit_issue_(rowIdx, h['操作'], '操作', 'NG', 'C-E06', '操作の値が不正です（「' + op + '」）'));
  }
  // C-E04: 時間変更/追加なのに時間タイプ空
  if ((op === '時間変更' || op === '追加') && !timeType) {
    issues.push(inputAudit_issue_(rowIdx, h['時間タイプ'], '時間タイプ', 'NG', 'C-E04', '操作が「' + op + '」ですが時間タイプが未設定です'));
  }
  // C-E05: 固定なのに時刻空
  if (timeType === '固定') {
    if (!startFixed) issues.push(inputAudit_issue_(rowIdx, h['開始時刻(固定)'], '開始時刻(固定)', 'NG', 'C-E05', '時間タイプが「固定」ですが開始時刻が未入力です'));
    if (!endFixed) issues.push(inputAudit_issue_(rowIdx, h['終了時刻(固定)'], '終了時刻(固定)', 'NG', 'C-E05', '時間タイプが「固定」ですが終了時刻が未入力です'));
  }
  // C-W01: 時間帯なのに希望最早/最遅空
  if (timeType === '時間帯') {
    if (!earliest || !latest) {
      issues.push(inputAudit_issue_(rowIdx, h['希望最早'], '希望最早', 'WARN', 'C-W01', '時間タイプが「時間帯」ですが希望最早/最遅が未入力です'));
    }
  }

  return issues;
}

// ============================================================
// スタッフ個別変更リクエスト判定
// ============================================================

function inputAudit_judgeStaffChangeRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var sid = inputAudit_getVal_(row, h['staff_id']);
  var date = inputAudit_getVal_(row, h['日付']);
  var rType = inputAudit_getVal_(row, h['制限タイプ']);
  var startTime = inputAudit_getVal_(row, h['開始時刻']);
  var endTime = inputAudit_getVal_(row, h['終了時刻']);

  // SC-E01
  if (!sid) issues.push(inputAudit_issue_(rowIdx, h['staff_id'], 'staff_id', 'NG', 'SC-E01', 'スタッフIDが未設定です'));
  // SC-E02
  if (!date) issues.push(inputAudit_issue_(rowIdx, h['日付'], '日付', 'NG', 'SC-E02', '日付が未入力です'));
  // SC-E03
  if (!rType) issues.push(inputAudit_issue_(rowIdx, h['制限タイプ'], '制限タイプ', 'NG', 'SC-E03', '制限タイプが未選択です'));
  // SC-E07: 制限タイプ不正
  if (rType && !inputAudit_isValidChoice(rType, '制限タイプ')) {
    issues.push(inputAudit_issue_(rowIdx, h['制限タイプ'], '制限タイプ', 'NG', 'SC-E07', '制限タイプの値が不正です（「' + rType + '」）'));
  }
  // SC-E04: 時間指定なのに時刻空
  if (rType === '時間指定') {
    if (!startTime) issues.push(inputAudit_issue_(rowIdx, h['開始時刻'], '開始時刻', 'NG', 'SC-E04', '制限タイプが「時間指定」ですが開始時刻が未入力です'));
    if (!endTime) issues.push(inputAudit_issue_(rowIdx, h['終了時刻'], '終了時刻', 'NG', 'SC-E04', '制限タイプが「時間指定」ですが終了時刻が未入力です'));
  }
  // SC-E05: 遅刻なのに開始空
  if (rType === '遅刻' && !startTime) {
    issues.push(inputAudit_issue_(rowIdx, h['開始時刻'], '開始時刻', 'NG', 'SC-E05', '制限タイプが「遅刻」ですが開始時刻が未入力です'));
  }
  // SC-E06: 早退なのに終了空
  if (rType === '早退' && !endTime) {
    issues.push(inputAudit_issue_(rowIdx, h['終了時刻'], '終了時刻', 'NG', 'SC-E06', '制限タイプが「早退」ですが終了時刻が未入力です'));
  }

  return issues;
}

// ============================================================
// イベントリクエスト判定
// ============================================================

function inputAudit_judgeEventRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var sid = inputAudit_getVal_(row, h['staff_id']);
  var date = inputAudit_getVal_(row, h['日付']);
  var title = inputAudit_getVal_(row, h['タイトル']);
  var timeMode = inputAudit_getVal_(row, h['時間指定方法']);
  var startTime = inputAudit_getVal_(row, h['開始時刻']);
  var endTime = inputAudit_getVal_(row, h['終了時刻']);
  var duration = inputAudit_getVal_(row, h['所要時間(分)']);
  var address = inputAudit_getVal_(row, h['住所']);
  var eventType = inputAudit_getVal_(row, h['イベント種別']);

  // EV-E01
  if (!sid) issues.push(inputAudit_issue_(rowIdx, h['staff_id'], 'staff_id', 'NG', 'EV-E01', 'スタッフIDが未設定です'));
  // EV-E02
  if (!date) issues.push(inputAudit_issue_(rowIdx, h['日付'], '日付', 'NG', 'EV-E02', '日付が未入力です'));
  // EV-E03
  if (!title) issues.push(inputAudit_issue_(rowIdx, h['タイトル'], 'タイトル', 'NG', 'EV-E03', 'タイトルが未入力です'));
  // EV-E06
  if (!timeMode) issues.push(inputAudit_issue_(rowIdx, h['時間指定方法'], '時間指定方法', 'NG', 'EV-E06', '時間指定方法が未選択です'));
  // EV-E04: 時間指定なのに時刻空
  if (timeMode === '時間指定') {
    if (!startTime) issues.push(inputAudit_issue_(rowIdx, h['開始時刻'], '開始時刻', 'NG', 'EV-E04', '時間指定方法が「時間指定」ですが開始時刻が未入力です'));
    if (!endTime) issues.push(inputAudit_issue_(rowIdx, h['終了時刻'], '終了時刻', 'NG', 'EV-E04', '時間指定方法が「時間指定」ですが終了時刻が未入力です'));
  }
  // EV-E05: 午前/午後/終日なのに所要時間空
  if ((timeMode === '午前' || timeMode === '午後' || timeMode === '終日') && !duration) {
    issues.push(inputAudit_issue_(rowIdx, h['所要時間(分)'], '所要時間(分)', 'NG', 'EV-E05', '時間指定方法が「' + timeMode + '」ですが所要時間が未入力です'));
  }
  // EV-W01
  if (!address) issues.push(inputAudit_issue_(rowIdx, h['住所'], '住所', 'WARN', 'EV-W01', '住所が未設定です。移動時間の計算に影響します'));
  // EV-W02
  if (!eventType) issues.push(inputAudit_issue_(rowIdx, h['イベント種別'], 'イベント種別', 'WARN', 'EV-W02', 'イベント種別が未選択です'));

  return issues;
}

// ============================================================
// スタッフ同行割付判定
// ============================================================

function inputAudit_judgeMentorRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var trainee = inputAudit_getVal_(row, h['trainee_staff_id']);
  var mentor = inputAudit_getVal_(row, h['mentor_staff_id']);
  var startDate = inputAudit_getVal_(row, h['開始日']);
  var endDate = inputAudit_getVal_(row, h['終了日']);

  // M-E01
  if (!trainee) issues.push(inputAudit_issue_(rowIdx, h['trainee_staff_id'], 'trainee_staff_id', 'NG', 'M-E01', '研修スタッフIDが未設定です'));
  // M-E02
  if (!mentor) issues.push(inputAudit_issue_(rowIdx, h['mentor_staff_id'], 'mentor_staff_id', 'NG', 'M-E02', 'メンタースタッフIDが未設定です'));
  // M-E03
  if (!startDate) issues.push(inputAudit_issue_(rowIdx, h['開始日'], '開始日', 'NG', 'M-E03', '開始日が未入力です'));
  // M-E04
  if (!endDate) issues.push(inputAudit_issue_(rowIdx, h['終了日'], '終了日', 'NG', 'M-E04', '終了日が未入力です'));
  // M-E05: 開始 > 終了
  if (startDate && endDate) {
    var sd = audit_parseDate_(startDate);
    var ed = audit_parseDate_(endDate);
    if (sd && ed && sd.getTime() > ed.getTime()) {
      issues.push(inputAudit_issue_(rowIdx, h['開始日'], '開始日', 'NG', 'M-E05', '開始日が終了日より後になっています'));
    }
  }
  // M-E06: 同一人物
  if (trainee && mentor && trainee === mentor) {
    issues.push(inputAudit_issue_(rowIdx, h['trainee_staff_id'], 'trainee_staff_id', 'NG', 'M-E06', '研修スタッフとメンターが同一人物です'));
  }

  return issues;
}

// ============================================================
// 週間訪問パターン判定
// ============================================================

function inputAudit_judgePatternRow_(headers, row, rowIdx) {
  var issues = [];
  var h = inputAudit_headerMap_(headers);

  var pid = inputAudit_getVal_(row, h['patient_id']);
  var dayCode = inputAudit_getVal_(row, h['曜日コード']);
  var startTime = inputAudit_getVal_(row, h['開始時刻']);
  var svcTime = inputAudit_getVal_(row, h['サービス時間']);
  var timeType = inputAudit_getVal_(row, h['時間タイプ']);

  // WP-E01
  if (!pid) issues.push(inputAudit_issue_(rowIdx, h['patient_id'], 'patient_id', 'NG', 'WP-E01', '患者IDが未設定です'));
  // WP-E02
  if (!dayCode) issues.push(inputAudit_issue_(rowIdx, h['曜日コード'], '曜日コード', 'NG', 'WP-E02', '曜日コードが未設定です'));
  // WP-E04
  if (!svcTime) issues.push(inputAudit_issue_(rowIdx, h['サービス時間'], 'サービス時間', 'NG', 'WP-E04', 'サービス時間が未入力です'));
  // WP-E05: 時間タイプの値が不正
  if (timeType && !inputAudit_isValidChoice(timeType, '時間タイプ')) {
    issues.push(inputAudit_issue_(rowIdx, h['時間タイプ'], '時間タイプ', 'NG', 'WP-E05', '時間タイプの値が不正です（「' + timeType + '」）'));
  }
  // WP-E03: 固定/時間帯なのに開始時刻空
  if ((timeType === '固定' || timeType === '時間帯') && !startTime) {
    issues.push(inputAudit_issue_(rowIdx, h['開始時刻'], '開始時刻', 'NG', 'WP-E03', '時間タイプが「' + timeType + '」ですが開始時刻が未入力です'));
  }
  // WP-W01: 時間タイプ空
  if (!timeType) issues.push(inputAudit_issue_(rowIdx, h['時間タイプ'], '時間タイプ', 'WARN', 'WP-W01', '時間タイプが未設定です。「時間帯」として扱われます'));
  // WP-W02: 午前/午後/終日なのに開始時刻あり
  if ((timeType === '午前' || timeType === '午後' || timeType === '終日') && startTime) {
    issues.push(inputAudit_issue_(rowIdx, h['開始時刻'], '開始時刻', 'WARN', 'WP-W02', '時間タイプが「' + timeType + '」ですが開始時刻が入力されています。開始時刻は無視されます'));
  }

  return issues;
}

// ============================================================
// ヘルパー関数
// ============================================================

/**
 * スプレッドシートからデータ取得
 */
function inputAudit_getSheetData_(sheetName) {
  var ssId = PropertiesService.getScriptProperties().getProperty('SS_ID');
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { headers: data[0] || [], rows: [] };
  return { headers: data[0], rows: data.slice(1) };
}

/**
 * ヘッダーマップを構築（名前→インデックス）
 */
function inputAudit_headerMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (h) map[h] = i;
  }
  // 部分一致で補完（audit_findHeaderIdx_ を使用）
  var targets = ['patient_id', '患者名', '稼働状況', '週訪問回数', '希望曜日', 'サービス時間',
    '時間タイプ', '希望時間帯（開始）', '希望時間帯（終了）', '緯度', '経度', '性別制限',
    '必要スタッフ数', '指定スタッフID', '指定タイプ', 'NGスタッフID', '継続希望', '曜日NG',
    'staff_id', 'スタッフ名', '性別', 'シフト開始', 'シフト終了', '勤務曜日', '最大訪問件数/日', '割当量',
    '日付', '操作', '開始時刻(固定)', '終了時刻(固定)', '希望最早', '希望最遅',
    '制限タイプ', '開始時刻', '終了時刻',
    'タイトル', '時間指定方法', '所要時間(分)', '住所', 'イベント種別',
    'trainee_staff_id', 'mentor_staff_id', '開始日', '終了日'];
  for (var j = 0; j < targets.length; j++) {
    var t = targets[j];
    if (!(t in map)) {
      var idx = audit_findHeaderIdx_(headers, t);
      if (idx >= 0) map[t] = idx;
    }
  }
  return map;
}

/**
 * セル値取得
 */
function inputAudit_getVal_(row, idx) {
  if (idx === undefined || idx === null || idx < 0 || idx >= row.length) return '';
  var v = row[idx];
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v;
  return String(v).trim();
}

/**
 * Issue オブジェクト生成
 */
function inputAudit_issue_(rowIdx, colIdx, field, level, ruleId, message) {
  return {
    rowIdx: rowIdx,
    colIdx: (colIdx !== undefined && colIdx !== null && colIdx >= 0) ? colIdx : -1,
    field: field,
    level: level,
    ruleId: ruleId,
    message: message
  };
}

/**
 * 空行判定
 */
function inputAudit_isEmptyRow_(row) {
  for (var i = 0; i < row.length; i++) {
    var v = row[i];
    if (v !== null && v !== undefined && String(v).trim() !== '') return false;
  }
  return true;
}
