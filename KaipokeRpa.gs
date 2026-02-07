// ==========================================
// カイポケ自動化 API連携
// ==========================================

// APIサーバーのベースURL（本番環境に合わせて変更）
var API_BASE_URL = "https://kaipoke-api.net";

// Google DriveのフォルダID（共有フォルダのID）
var DRIVE_FOLDER_ID = "1ABCxxxxxxxxxxxxxxxx";

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

  var payload = {
    "month": month || getCurrentMonth(),
    "upload_to_drive": true,
    "drive_folder_id": DRIVE_FOLDER_ID
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
      var msg = "CSV出力完了!\n出力先: " + result.result.file_path;
      if (result.result.drive_file_id) {
        msg += "\nGoogle Driveにアップロード済み";
      }
      return {
        "success": true,
        "message": msg,
        "data": result.result
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
// 差分確認（プレビュー）
// ==========================================
function checkDiff(month, weekStart) {
  var url = API_BASE_URL + "/api/diff";

  var monthStr = (month || getCurrentMonth()).replace("-", "");
  var payload = {
    "month": month || getCurrentMonth(),
    "week_start": weekStart,
    "current_csv": "data/current_" + monthStr + ".csv",
    "optimized_csv": "data/optimized_" + monthStr + ".csv"
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
        "message": "差分確認結果:\n" +
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
// 差分適用
// ==========================================
function runApply(month, weekStart) {
  var url = API_BASE_URL + "/api/apply";

  var monthStr = (month || getCurrentMonth()).replace("-", "");
  var payload = {
    "month": month || getCurrentMonth(),
    "week_start": weekStart,
    "current_csv": "data/current_" + monthStr + ".csv",
    "optimized_csv": "data/optimized_" + monthStr + ".csv"
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
        "message": "差分適用完了!\n" +
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
    var fileName = "optimized_" + monthStr + ".csv";

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
 * 管理者シートの任意の行で、A列に「カイポケパスワード」と書いて
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
      if (cellA === 'カイポケパスワード') {
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
