/**
 * 監査（合否判定）ビュー - 判定エンジン
 *
 * Expected（期待）とActual（実績）を比較して○/△/×を判定
 */

// ============================================================
// 判定メイン
// ============================================================

/**
 * 患者×日の判定を実行
 * @param {Object} dataset - 週データセット
 * @param {Object} expectedByPidDate - 期待Map
 * @param {string} pid - 患者ID
 * @param {string} dateStr - 日付
 * @return {Object} { status, tags, checks }
 */
function audit_judgePatientDay_(dataset, expectedByPidDate, pid, dateStr) {
  var key = audit_makePdKey_(pid, dateStr);
  var dayExpected = expectedByPidDate[key];
  var actuals = dataset.actualPlanMap[key] || [];
  var events = dataset.eventMap[key] || [];
  var bufferMin = AUDIT_CONFIG.TIME_BUFFER_MIN;

  var checks = [];
  var tags = [];
  var status = AUDIT_STATUS.OK;

  // 期待がない場合
  if (!dayExpected || dayExpected.visits.length === 0) {
    if (actuals.length > 0) {
      // 期待がないのに実績がある → EXTRA_ACTUAL (WARN)
      status = AUDIT_STATUS.WARN;
      tags.push(AUDIT_TAGS.EXTRA_ACTUAL);
      checks.push({
        type: 'extra',
        status: AUDIT_STATUS.WARN,
        reason: '期待がないのに実績があります'
      });
    }
    return { status: status, tags: tags, checks: checks };
  }

  var expectedVisits = dayExpected.visits;

  // === キャンセル判定 ===
  var cancelledExpects = expectedVisits.filter(function(e) { return e.isCancelled; });
  if (cancelledExpects.length > 0 && actuals.length > 0) {
    // キャンセルなのに実績がある → NG
    status = AUDIT_STATUS.NG;
    tags.push(AUDIT_TAGS.CANCELLED_BUT_VISITED);
    checks.push({
      type: 'cancel',
      status: AUDIT_STATUS.NG,
      reason: 'キャンセル期待だが実績があります'
    });
    return { status: status, tags: tags, checks: checks };
  }

  if (cancelledExpects.length > 0 && actuals.length === 0) {
    // キャンセル期待で実績なし → OK
    tags.push('CHANGE(キャンセル)');
    checks.push({
      type: 'cancel',
      status: AUDIT_STATUS.OK,
      reason: 'キャンセル期待通り'
    });
    return { status: status, tags: tags, checks: checks };
  }

  // === イベント衝突判定 ===
  if (events.length > 0 && actuals.length > 0) {
    var conflictEvent = null;
    var conflictActual = null;

    actuals.forEach(function(actual) {
      events.forEach(function(event) {
        if (audit_isTimeOverlap_(actual.startMin, actual.endMin, event.startMin, event.endMin)) {
          conflictEvent = event;
          conflictActual = actual;
        }
      });
    });

    if (conflictEvent) {
      status = AUDIT_STATUS.NG;
      tags.push(AUDIT_TAGS.EVENT_CONFLICT);
      checks.push({
        type: 'eventConflict',
        status: AUDIT_STATUS.NG,
        event: conflictEvent,
        actual: conflictActual,
        reason: 'イベント(' + (conflictEvent.title || conflictEvent.eventId) + ')と時間が重複'
      });
      // イベント衝突があっても他のチェックも続行
    }
  }

  // === 未割当判定 ===
  var unassignedActuals = actuals.filter(function(a) { return a.isUnassigned; });
  if (unassignedActuals.length > 0) {
    status = AUDIT_STATUS.NG;
    tags.push(AUDIT_TAGS.UNASSIGNED);
    checks.push({
      type: 'unassigned',
      status: AUDIT_STATUS.NG,
      reason: '未割当の実績があります'
    });
  }

  // === 期待と実績の突合 ===
  var activeExpects = expectedVisits.filter(function(e) { return !e.isCancelled; });
  var usedActualIdxs = {};

  activeExpects.forEach(function(exp) {
    // 最も近い実績を探す
    var bestActualIdx = -1;
    var bestDiff = Infinity;

    actuals.forEach(function(actual, idx) {
      if (usedActualIdxs[idx]) return;

      var diff = audit_calcTimeDiff_(exp, actual);
      if (diff < bestDiff) {
        bestDiff = diff;
        bestActualIdx = idx;
      }
    });

    if (bestActualIdx >= 0) {
      usedActualIdxs[bestActualIdx] = true;
      var actual = actuals[bestActualIdx];

      // 時間判定
      var timeCheck = audit_checkTimeWindow_(exp, actual, bufferMin);
      checks.push(timeCheck);

      if (timeCheck.status === AUDIT_STATUS.NG) {
        status = AUDIT_STATUS.NG;
        tags.push(AUDIT_TAGS.OUT_OF_WINDOW);
      } else if (timeCheck.status === AUDIT_STATUS.WARN && status !== AUDIT_STATUS.NG) {
        status = AUDIT_STATUS.WARN;
        if (timeCheck.diffMin !== 0) {
          var diffStr = (timeCheck.diffMin > 0 ? '+' : '') + timeCheck.diffMin + 'm';
          tags.push(AUDIT_TAGS.TIME_DIFF + '(' + diffStr + ')');
        }
      }
    } else {
      // 期待に対応する実績がない → MISSING_ACTUAL
      if (status !== AUDIT_STATUS.NG) {
        status = AUDIT_STATUS.WARN;
      }
      tags.push(AUDIT_TAGS.MISSING_ACTUAL);
      checks.push({
        type: 'missing',
        status: AUDIT_STATUS.WARN,
        expected: exp,
        reason: '期待に対応する実績がありません'
      });
    }

    // ソースタグを追加
    if (tags.indexOf(exp.source) < 0 && tags.indexOf(AUDIT_TAGS.OUT_OF_WINDOW) < 0) {
      var sourceTag = audit_getSourceTag_(exp);
      if (sourceTag && tags.indexOf(sourceTag) < 0) {
        tags.push(sourceTag);
      }
    }
  });

  // 使われなかった実績（余剰）
  actuals.forEach(function(actual, idx) {
    if (!usedActualIdxs[idx]) {
      if (status !== AUDIT_STATUS.NG) {
        status = AUDIT_STATUS.WARN;
      }
      tags.push(AUDIT_TAGS.EXTRA_ACTUAL);
      checks.push({
        type: 'extra',
        status: AUDIT_STATUS.WARN,
        actual: actual,
        reason: '余剰の実績があります'
      });
    }
  });

  // タグの重複除去
  tags = tags.filter(function(tag, idx, arr) {
    return arr.indexOf(tag) === idx;
  });

  return { status: status, tags: tags, checks: checks };
}

/**
 * 時間窓チェック
 */
function audit_checkTimeWindow_(expected, actual, bufferMin) {
  var timeType = expected.timeType || '時間帯';
  var result = {
    type: 'timeWindow',
    status: AUDIT_STATUS.OK,
    expected: {
      timeType: timeType,
      startMin: expected.startMin,
      endMin: expected.endMin,
      earliestMin: expected.earliestMin,
      latestMin: expected.latestMin
    },
    actual: {
      startMin: actual.startMin,
      endMin: actual.endMin
    },
    diffMin: 0,
    reason: ''
  };

  if (actual.startMin === null || actual.endMin === null) {
    result.status = AUDIT_STATUS.WARN;
    result.reason = '実績の時刻が不明';
    return result;
  }

  if (timeType === '固定') {
    // 固定: 開始時刻のずれで判定
    if (expected.startMin !== null) {
      var diff = actual.startMin - expected.startMin;
      result.diffMin = diff;

      if (Math.abs(diff) <= bufferMin) {
        result.status = (diff === 0) ? AUDIT_STATUS.OK : AUDIT_STATUS.WARN;
        result.reason = (diff === 0) ? '' : '開始時刻が' + diff + '分ずれ';
      } else {
        result.status = AUDIT_STATUS.NG;
        result.reason = '開始時刻が' + diff + '分ずれ（バッファ超過）';
      }
    }
  } else {
    // 時間帯/午前/午後/終日: 範囲内かどうかで判定
    var earliest = expected.earliestMin;
    var latest = expected.latestMin;

    if (earliest === null && latest === null) {
      // 範囲が不明 → OK扱い
      return result;
    }

    var startOk = (earliest === null || actual.startMin >= earliest - bufferMin);
    var endOk = (latest === null || actual.endMin <= latest + bufferMin);

    if (earliest !== null && actual.startMin < earliest) {
      result.diffMin = actual.startMin - earliest;
    } else if (latest !== null && actual.endMin > latest) {
      result.diffMin = actual.endMin - latest;
    }

    var startInRange = (earliest === null || actual.startMin >= earliest);
    var endInRange = (latest === null || actual.endMin <= latest);

    if (startInRange && endInRange) {
      result.status = AUDIT_STATUS.OK;
    } else if (startOk && endOk) {
      result.status = AUDIT_STATUS.WARN;
      result.reason = '期待範囲外だがバッファ内';
    } else {
      result.status = AUDIT_STATUS.NG;
      result.reason = '期待範囲外（バッファ超過）';
    }
  }

  return result;
}

/**
 * 時刻差分を計算（突合用）
 */
function audit_calcTimeDiff_(expected, actual) {
  var expStart = expected.startMin !== null ? expected.startMin : expected.earliestMin;
  var actStart = actual.startMin;

  if (expStart === null || actStart === null) return 9999;

  return Math.abs(actStart - expStart);
}

/**
 * ソースからタグを取得
 */
function audit_getSourceTag_(exp) {
  if (exp.source === AUDIT_SOURCE.MASTER) return AUDIT_TAGS.MASTER;
  if (exp.source === AUDIT_SOURCE.CHANGE) {
    if (exp.op === '時間変更') return AUDIT_TAGS.CHANGE_TIME;
    if (exp.op === '追加') return AUDIT_TAGS.CHANGE_ADD;
  }
  if (exp.source === AUDIT_SOURCE.SPECIAL_ADD) return AUDIT_TAGS.SPECIAL_ADD;
  if (exp.source === AUDIT_SOURCE.SPECIAL_REPLACE) return AUDIT_TAGS.SPECIAL_REPLACE;
  return null;
}

// ============================================================
// データセット拡張（Expected付き）
// ============================================================

/**
 * データセットにExpectedを付与
 */
function audit_enrichDatasetWithExpected_(dataset) {
  var expectedByPidDate = audit_buildExpected_(dataset);
  dataset.expectedByPidDate = expectedByPidDate;
  return dataset;
}
