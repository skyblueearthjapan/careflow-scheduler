// ==========================================
// カイポケ自動化 API連携
// ==========================================

// APIサーバーのベースURL（本番環境に合わせて変更）
var API_BASE_URL = "https://kaipoke-api.net";

// Google DriveのフォルダID（共有フォルダのID）
var DRIVE_FOLDER_ID = "1tQJKZDjonFwiY6wYYgx1iVgu4cM98vRp";

// ==========================================
// サーバー状態確認
// ==========================================
function checkServerStatus() {
  var url = API_BASE_URL + "/api/status";

  var options = {
    "method": "get",
    "muteHttpExceptions": true,
    "headers": {
      "Content-Type": "application/json"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());

    if (result.status === "running") {
      if (result.current_task && result.current_task.running) {
        return {
          "status": "busy",
          "message": "タスク実行中: " + result.current_task.command
        };
      } else {
        return {
          "status": "ready",
          "message": "サーバー稼働中（待機中）"
        };
      }
    }
    return {
      "status": "unknown",
      "message": "サーバー状態不明"
    };
  } catch (e) {
    return {
      "status": "error",
      "message": "サーバーに接続できません: " + e.message
    };
  }
}

// ==========================================
// 月間スケジュール展開
// ==========================================
function runExpand(month) {
  var url = API_BASE_URL + "/api/expand";

  var payload = {
    "month": month || getCurrentMonth()
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.success) {
      var r = result.result;
      return {
        "success": true,
        "message": "展開完了!\n" +
                   "成功: " + r.success + "件\n" +
                   "スキップ: " + r.skipped + "件\n" +
                   "失敗: " + r.failed + "件\n" +
                   "合計: " + r.total + "件",
        "data": r
      };
    } else if (statusCode === 409) {
      return {
        "success": false,
        "message": "エラー: " + result.error
      };
    } else {
      return {
        "success": false,
        "message": "エラー: " + (result.error || "不明なエラー")
      };
    }
  } catch (e) {
    return {
      "success": false,
      "message": "サーバー接続エラー: " + e.message
    };
  }
}

// ==========================================
// CSV出力（Google Driveアップロード付き）
// ==========================================
function runExport(month) {
  var url = API_BASE_URL + "/api/export";
  var targetMonth = month || getCurrentMonth();

  var payload = {
    "month": targetMonth
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.success) {
      var csvContent = result.result.csv_content;
      var driveResult = null;

      // csv_contentがあればGoogle Driveに保存
      if (csvContent) {
        driveResult = saveCsvToDrive_(csvContent, targetMonth);
      }

      var msg = "CSV出力完了!";
      if (result.result.file_path) {
        msg += "\nVPS出力先: " + result.result.file_path;
      }
      if (driveResult && driveResult.success) {
        msg += "\nGoogle Drive保存: " + driveResult.fileName;
      } else if (csvContent && driveResult && !driveResult.success) {
        msg += "\nDrive保存エラー: " + driveResult.message;
      }

      return {
        "success": true,
        "message": msg,
        "data": result.result,
        "driveFileId": driveResult ? driveResult.fileId : null
      };
    } else if (statusCode === 409) {
      return {
        "success": false,
        "message": "エラー: " + result.error
      };
    } else {
      return {
        "success": false,
        "message": "エラー: " + (result.error || "不明なエラー")
      };
    }
  } catch (e) {
    return {
      "success": false,
      "message": "サーバー接続エラー: " + e.message
    };
  }
}

/**
 * CSVテキストをGoogle Driveに保存
 * @param {string} csvContent - CSV文字列
 * @param {string} month - 対象月（YYYY-MM形式）
 * @return {Object} {success, fileId, fileName, message}
 */
function saveCsvToDrive_(csvContent, month) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var monthStr = month.replace("-", "");
    var fileName = "kaipoke_current_" + monthStr + ".csv";

    // 既存ファイルがあれば削除（上書き）
    var existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    // BOM付きUTF-8で保存（Excel対応）
    var bom = "\uFEFF";
    var blob = Utilities.newBlob(bom + csvContent, "text/csv", fileName);
    var file = folder.createFile(blob);

    return {
      "success": true,
      "fileId": file.getId(),
      "fileName": fileName,
      "message": "保存完了"
    };
  } catch (e) {
    return {
      "success": false,
      "fileId": null,
      "fileName": null,
      "message": e.message
    };
  }
}

// ==========================================
// 差分確認（プレビュー）- パターンB: CSV内容を直接送信
// ==========================================
function checkDiff(month, weekStart) {
  var url = API_BASE_URL + "/api/diff";
  var targetMonth = month || getCurrentMonth();
  var monthStr = targetMonth.replace("-", "");

  // weekStart必須チェック
  if (!weekStart) {
    return {
      "success": false,
      "message": "エラー: 対象週が指定されていません。"
    };
  }

  // 週範囲を算出（YYYYMMDD形式 + YYYY-MM-DD形式）
  var weekRange = getWeekRange_(weekStart);
  console.log('[checkDiff] month=' + targetMonth + ' weekStart=' + weekStart + ' weekEnd=' + weekRange.endDate);

  // --- Google DriveからCSVを読み込み ---
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // 1. kaipoke_current_YYYYMM.csv を読み込み
  var currentFileName = "kaipoke_current_" + monthStr + ".csv";
  var currentCsvContent = readCsvContentFromDrive_(folder, currentFileName);
  if (currentCsvContent === null) {
    return {
      "success": false,
      "message": "エラー: Google Driveに「" + currentFileName + "」が見つかりません。\n先にCSV出力（ステップ3）を実行してください。"
    };
  }
  console.log('[checkDiff] currentCSV読み込み完了: ' + currentFileName + ' (' + currentCsvContent.length + '文字)');

  // 2. gas_optimized_YYYYMMDD_YYYYMMDD.csv を読み込み
  var optimizedFileName = "gas_optimized_" + weekRange.startStr + "_" + weekRange.endStr + ".csv";
  var optimizedCsvContent = readCsvContentFromDrive_(folder, optimizedFileName);
  if (optimizedCsvContent === null) {
    return {
      "success": false,
      "message": "エラー: Google Driveに「" + optimizedFileName + "」が見つかりません。\n先にGAS側のCSV出力を実行してください。"
    };
  }
  console.log('[checkDiff] optimizedCSV読み込み完了: ' + optimizedFileName + ' (' + optimizedCsvContent.length + '文字)');

  // 3. CSV内容を直接送信
  var payload = {
    "month": targetMonth,
    "week_start": weekStart,
    "week_end": weekRange.endDate,
    "current_csv_content": currentCsvContent,
    "optimized_csv_content": optimizedCsvContent
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.success) {
      var d = result.diff;
      return {
        "success": true,
        "message": "差分確認結果（" + weekStart + " 〜 " + weekRange.endDate + "）:\n" +
                   "追加予定: " + d.add + "件\n" +
                   "削除予定: " + d.remove + "件\n" +
                   "変更予定: " + d.modify + "件",
        "data": d
      };
    } else {
      return {
        "success": false,
        "message": "エラー: " + (result.error || "不明なエラー")
      };
    }
  } catch (e) {
    return {
      "success": false,
      "message": "サーバー接続エラー: " + e.message
    };
  }
}

// ==========================================
// 差分適用 - パターンB: CSV内容を直接送信
// ==========================================
function runApply(month, weekStart) {
  var url = API_BASE_URL + "/api/apply";
  var targetMonth = month || getCurrentMonth();
  var monthStr = targetMonth.replace("-", "");

  // weekStart必須チェック
  if (!weekStart) {
    return {
      "success": false,
      "message": "エラー: 対象週が指定されていません。"
    };
  }

  // 週範囲を算出
  var weekRange = getWeekRange_(weekStart);
  console.log('[runApply] month=' + targetMonth + ' weekStart=' + weekStart + ' weekEnd=' + weekRange.endDate);

  // --- Google DriveからCSVを読み込み ---
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // 1. kaipoke_current_YYYYMM.csv を読み込み
  var currentFileName = "kaipoke_current_" + monthStr + ".csv";
  var currentCsvContent = readCsvContentFromDrive_(folder, currentFileName);
  if (currentCsvContent === null) {
    return {
      "success": false,
      "message": "エラー: Google Driveに「" + currentFileName + "」が見つかりません。\n先にCSV出力（ステップ3）を実行してください。"
    };
  }
  console.log('[runApply] currentCSV読み込み完了: ' + currentFileName);

  // 2. gas_optimized_YYYYMMDD_YYYYMMDD.csv を読み込み
  var optimizedFileName = "gas_optimized_" + weekRange.startStr + "_" + weekRange.endStr + ".csv";
  var optimizedCsvContent = readCsvContentFromDrive_(folder, optimizedFileName);
  if (optimizedCsvContent === null) {
    return {
      "success": false,
      "message": "エラー: Google Driveに「" + optimizedFileName + "」が見つかりません。\n先にGAS側のCSV出力を実行してください。"
    };
  }
  console.log('[runApply] optimizedCSV読み込み完了: ' + optimizedFileName);

  // 3. CSV内容を直接送信
  var payload = {
    "month": targetMonth,
    "week_start": weekStart,
    "week_end": weekRange.endDate,
    "current_csv_content": currentCsvContent,
    "optimized_csv_content": optimizedCsvContent
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.success) {
      var d = result.result.diff;
      var a = result.result.applied;
      return {
        "success": true,
        "message": "差分適用完了!（" + weekStart + " 〜 " + weekRange.endDate + "）\n" +
                   "追加: " + a.add + "/" + d.add + "件\n" +
                   "削除: " + a.remove + "/" + d.remove + "件\n" +
                   "変更: " + a.modify + "/" + d.modify + "件",
        "data": result.result
      };
    } else if (statusCode === 400) {
      return {
        "success": false,
        "message": "パラメータエラー: " + result.error
      };
    } else if (statusCode === 409) {
      return {
        "success": false,
        "message": "エラー: " + result.error
      };
    } else {
      return {
        "success": false,
        "message": "エラー: " + (result.error || "不明なエラー")
      };
    }
  } catch (e) {
    return {
      "success": false,
      "message": "サーバー接続エラー: " + e.message
    };
  }
}

// ==========================================
// 接続テスト
// ==========================================
function runConnectionTest() {
  var url = API_BASE_URL + "/api/test";

  var testPayload = {
    "action": "ping",
    "timestamp": new Date().toISOString(),
    "source": "gas_sidebar"
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(testPayload),
    "muteHttpExceptions": true
  };

  try {
    var startTime = new Date().getTime();
    var response = UrlFetchApp.fetch(url, options);
    var endTime = new Date().getTime();
    var responseTime = endTime - startTime;

    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.success) {
      return {
        "success": true,
        "message": "接続テスト成功!\n" +
                   "ステータス: OK\n" +
                   "応答時間: " + responseTime + "ms\n" +
                   "サーバー時刻: " + (result.server_time || "N/A") + "\n" +
                   "メッセージ: " + (result.message || "テスト完了")
      };
    } else {
      return {
        "success": false,
        "message": "接続テスト失敗\n" +
                   "ステータスコード: " + statusCode + "\n" +
                   "エラー: " + (result.error || "不明なエラー")
      };
    }
  } catch (e) {
    // /api/test が存在しない場合は /api/status にフォールバック
    try {
      var statusUrl = API_BASE_URL + "/api/status";
      var statusOptions = {
        "method": "get",
        "muteHttpExceptions": true
      };

      var startTime2 = new Date().getTime();
      var statusResponse = UrlFetchApp.fetch(statusUrl, statusOptions);
      var endTime2 = new Date().getTime();
      var responseTime2 = endTime2 - startTime2;

      var statusResult = JSON.parse(statusResponse.getContentText());

      if (statusResult.status === "running") {
        return {
          "success": true,
          "message": "接続テスト成功!\n" +
                     "（/api/status で確認）\n" +
                     "ステータス: " + statusResult.status + "\n" +
                     "応答時間: " + responseTime2 + "ms"
        };
      } else {
        return {
          "success": false,
          "message": "サーバー状態が異常です: " + statusResult.status
        };
      }
    } catch (e2) {
      return {
        "success": false,
        "message": "サーバーに接続できません\n" +
                   "URL: " + API_BASE_URL + "\n" +
                   "エラー: " + e.message
      };
    }
  }
}

// ==========================================
// ログ取得
// ==========================================
function kaipoke_logs(tail) {
  var url = API_BASE_URL + "/api/kaipoke/logs";
  if (tail) {
    url += "?tail=" + tail;
  }

  var options = {
    "method": "get",
    "muteHttpExceptions": true,
    "headers": {
      "Content-Type": "application/json"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.ok) {
      return {
        "success": true,
        "lines": result.lines || []
      };
    } else {
      return {
        "success": false,
        "lines": [],
        "message": result.error || "ログ取得エラー"
      };
    }
  } catch (e) {
    return {
      "success": false,
      "lines": ["[ローカル] ログ取得失敗: " + e.message],
      "message": e.message
    };
  }
}

// ==========================================
// VNC URL取得
// ==========================================
function kaipoke_vncUrl() {
  var url = API_BASE_URL + "/api/kaipoke/vnc-url";

  var options = {
    "method": "get",
    "muteHttpExceptions": true,
    "headers": {
      "Content-Type": "application/json"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.ok) {
      return {
        "success": true,
        "url": result.url,
        "ready": result.ready || false
      };
    } else {
      return {
        "success": false,
        "url": null,
        "message": result.error || "VNC URL取得エラー"
      };
    }
  } catch (e) {
    return {
      "success": false,
      "url": null,
      "message": e.message
    };
  }
}

// ==========================================
// 拡張ステータス取得（VNC URL含む）
// ==========================================
function kaipoke_status() {
  var url = API_BASE_URL + "/api/kaipoke/status";

  var options = {
    "method": "get",
    "muteHttpExceptions": true,
    "headers": {
      "Content-Type": "application/json"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var result = JSON.parse(response.getContentText());

    if (statusCode === 200 && result.ok) {
      return {
        "success": true,
        "server": result.server || {},
        "job": result.job || {},
        "vnc": result.vnc || {},
        "message": result.message || "OK"
      };
    } else {
      return {
        "success": false,
        "message": result.error || "ステータス取得エラー"
      };
    }
  } catch (e) {
    return {
      "success": false,
      "message": "サーバー接続エラー: " + e.message
    };
  }
}

// ==========================================
// 設定更新（DriveフォルダID）
// ==========================================
function setDriveFolderId(folderId) {
  var url = API_BASE_URL + "/api/config";

  var payload = {
    "folder_id": folderId
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = JSON.parse(response.getContentText());

    if (result.success) {
      return {
        "success": true,
        "message": "設定を更新しました"
      };
    } else {
      return {
        "success": false,
        "message": "エラー: " + result.error
      };
    }
  } catch (e) {
    return {
      "success": false,
      "message": "接続エラー: " + e.message
    };
  }
}

// ==========================================
// Google Drive操作（最適化CSV保存）
// ==========================================

/**
 * 最適化結果をCSVとしてGoogle Driveに保存
 * @param {Array} data - 2次元配列のスケジュールデータ
 * @param {string} month - 対象月（YYYY-MM形式）
 * @return {Object} 結果
 */
function saveOptimizedCsvToDrive(data, month) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var monthStr = month.replace("-", "");
    var fileName = "gas_optimized_" + monthStr + ".csv";

    // ヘッダー行
    var headers = [
      "職員名1", "職種1", "職員名2", "職種2", "同行2",
      "職員名3", "職種3", "同行3", "事業所名",
      "日付", "曜日", "利用者", "業務種別", "サービス内容",
      "開始時間", "終了時間", "提供時間", "備考"
    ];

    // CSVコンテンツを生成
    var csvContent = headers.join(",") + "\n";
    for (var i = 0; i < data.length; i++) {
      var row = data[i].map(function(cell) {
        // カンマや改行を含む場合はダブルクォートで囲む
        if (cell && (cell.toString().indexOf(",") >= 0 || cell.toString().indexOf("\n") >= 0)) {
          return '"' + cell.toString().replace(/"/g, '""') + '"';
        }
        return cell || "";
      });
      csvContent += row.join(",") + "\n";
    }

    // 既存ファイルがあれば削除
    var existingFiles = folder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    // 新規作成
    var blob = Utilities.newBlob(csvContent, "text/csv", fileName);
    var file = folder.createFile(blob);

    return {
      "success": true,
      "message": "保存完了: " + fileName,
      "fileId": file.getId()
    };
  } catch (e) {
    return {
      "success": false,
      "message": "保存エラー: " + e.message
    };
  }
}

/**
 * Google DriveからCSVを読み込み
 * @param {string} fileName - ファイル名
 * @return {Array} 2次元配列
 */
function loadCsvFromDrive(fileName) {
  try {
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      throw new Error("ファイルが見つかりません: " + fileName);
    }

    var file = files.next();
    var content = file.getBlob().getDataAsString("UTF-8");

    // BOM除去
    if (content.charCodeAt(0) === 0xFEFF) {
      content = content.substring(1);
    }

    // CSV解析
    var rows = Utilities.parseCsv(content);
    return rows;
  } catch (e) {
    throw new Error("CSV読み込みエラー: " + e.message);
  }
}

// ==========================================
// ユーティリティ関数
// ==========================================

/**
 * Google DriveフォルダからCSVファイルの内容を読み込む
 * @param {Folder} folder - Google Driveフォルダ
 * @param {string} fileName - ファイル名
 * @return {string|null} CSV文字列（BOM除去済み）、ファイルが無い場合はnull
 */
function readCsvContentFromDrive_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    console.log('[readCsvContentFromDrive_] ファイルが見つかりません: ' + fileName);
    return null;
  }
  var file = files.next();
  var content = file.getBlob().getDataAsString("UTF-8");
  // BOM除去
  if (content.charCodeAt(0) === 0xFEFF) {
    content = content.substring(1);
  }
  return content;
}

/**
 * 週の開始日から週範囲を算出
 * @param {string} weekStartStr - 週開始日（YYYY-MM-DD形式）
 * @return {Object} { startStr: 'YYYYMMDD', endStr: 'YYYYMMDD', endDate: 'YYYY-MM-DD' }
 */
function getWeekRange_(weekStartStr) {
  var parts = weekStartStr.split("-");
  var ws = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  var we = new Date(ws);
  we.setDate(we.getDate() + 6);

  var startStr = Utilities.formatDate(ws, "JST", "yyyyMMdd");
  var endStr = Utilities.formatDate(we, "JST", "yyyyMMdd");
  var endDate = Utilities.formatDate(we, "JST", "yyyy-MM-dd");

  return {
    startStr: startStr,
    endStr: endStr,
    endDate: endDate
  };
}

/**
 * 現在の月をYYYY-MM形式で取得
 */
function getCurrentMonth() {
  var now = new Date();
  var year = now.getFullYear();
  var month = ("0" + (now.getMonth() + 1)).slice(-2);
  return year + "-" + month;
}

/**
 * 週の開始日（月曜日）を取得
 */
function getWeekStart(date) {
  var d = new Date(date);
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
}

// ==========================================
// パスワード認証
// ==========================================

/**
 * スプレッドシートの「管理者」シートからカイポケパスワードを取得
 * 管理者シートの任意の行で、A列に「パスワード」（または「カイポケパスワード」）と書いて
 * B列にパスワードを設定する
 * @returns {string|null} パスワード文字列、未設定の場合はnull
 */
function getKaipokePassword_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('管理者');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      var cellA = String(data[i][0] || '').trim();
      if (cellA === 'パスワード' || cellA === 'カイポケパスワード') {
        var pw = String(data[i][1] || '').trim();
        return pw || null;
      }
    }
    return null;
  } catch (e) {
    console.error('getKaipokePassword_ error:', e);
    return null;
  }
}

/**
 * カイポケ自動化のパスワードを検証
 * @param {string} inputPassword - ユーザーが入力したパスワード
 * @returns {Object} {success: boolean, message: string}
 */
function verifyKaipokePassword(inputPassword) {
  var storedPassword = getKaipokePassword_();

  if (!storedPassword) {
    // パスワード未設定の場合はアクセスを許可
    return { success: true, message: 'パスワード未設定のためアクセス許可' };
  }

  if (String(inputPassword || '').trim() === storedPassword) {
    return { success: true, message: '認証成功' };
  } else {
    return { success: false, message: 'パスワードが正しくありません' };
  }
}

// ==========================================
// インターロック設定
// ==========================================

/**
 * インターロック設定を取得
 * @returns {Object} {expandMinMonth, expandCompleted[], applyMinWeek, applyCompleted[]}
 */
function getInterlockSettings() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('管理者');
    if (!sheet) return { expandMinMonth: null, expandCompleted: [], applyMinWeek: null, applyCompleted: [] };

    var data = sheet.getDataRange().getValues();
    var settings = {
      expandMinMonth: null,
      expandCompleted: [],
      applyMinWeek: null,
      applyCompleted: []
    };

    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0] || '').trim();
      var val = String(data[i][1] || '').trim();

      if (key === '展開制限月') {
        settings.expandMinMonth = val || null;
      } else if (key === '展開完了月') {
        settings.expandCompleted = val ? val.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];
      } else if (key === '差分制限週') {
        settings.applyMinWeek = val || null;
      } else if (key === '差分完了週') {
        settings.applyCompleted = val ? val.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];
      }
    }

    return settings;
  } catch (e) {
    console.error('getInterlockSettings error:', e);
    return { expandMinMonth: null, expandCompleted: [], applyMinWeek: null, applyCompleted: [] };
  }
}

/**
 * インターロック設定を保存（パスワード必須）
 * @param {string} key - 設定キー（展開制限月/展開完了月/差分制限週/差分完了週）
 * @param {string} value - 設定値
 * @param {string} password - パスワード
 * @returns {Object} {success: boolean, message: string}
 */
function saveInterlockSetting(key, value, password) {
  var result = verifyKaipokePassword(password);
  if (!result.success) {
    return { success: false, message: 'パスワードが正しくありません' };
  }

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('管理者');
    if (!sheet) return { success: false, message: '管理者シートが見つかりません' };

    var data = sheet.getDataRange().getValues();
    var found = false;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }

    if (!found) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1).setValue(key);
      sheet.getRange(lastRow + 1, 2).setValue(value);
    }

    return { success: true, message: '設定を保存しました' };
  } catch (e) {
    return { success: false, message: '保存エラー: ' + e.message };
  }
}

/**
 * 展開完了月を自動追加（パスワード不要）
 */
function markExpandCompleted(month) {
  return addToCompletedList_('展開完了月', month);
}

/**
 * 差分適用完了週を自動追加（パスワード不要）
 */
function markApplyCompleted(weekStart) {
  return addToCompletedList_('差分完了週', weekStart);
}

/**
 * 完了リストに値を追加するヘルパー
 */
function addToCompletedList_(key, value) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('管理者');
    if (!sheet) return { success: false };

    var data = sheet.getDataRange().getValues();
    var found = false;

    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === key) {
        var existing = String(data[i][1] || '').trim();
        var items = existing ? existing.split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];
        if (items.indexOf(value) < 0) {
          items.push(value);
        }
        sheet.getRange(i + 1, 2).setValue(items.join(','));
        found = true;
        break;
      }
    }

    if (!found) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1).setValue(key);
      sheet.getRange(lastRow + 1, 2).setValue(value);
    }

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// サイドバー表示
// ==========================================

/**
 * カイポケ自動化サイドバーを表示
 */
function showKaipokeRpaSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('KaipokeRpaSidebar')
    .setTitle('カイポケ自動化')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}
