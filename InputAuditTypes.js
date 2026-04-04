/**
 * 入力データ監査 - 定数・ルール定義
 *
 * Python自動算出エンジンが必要とする項目に基づき、
 * 各入力シートのセルごとにNG/WARNを判定するためのルール定義。
 */

// ============================================================
// ステータス定数
// ============================================================

var INPUT_AUDIT_STATUS = {
  OK: 'OK',
  WARN: 'WARN',
  NG: 'NG'
};

// ============================================================
// 選択肢の正規値定義（バリデーション用）
// ============================================================

var INPUT_AUDIT_VALID_VALUES = {
  '稼働状況': ['稼働', '休止', '未開始', '未契約', '解約'],
  '時間タイプ': ['時間帯', '午前', '午後', '終日', '固定'],
  '性別制限': ['', '男性のみ', '女性のみ'],
  '性別': ['', '男性', '女性'],
  '継続希望': ['同じ人希望', 'ローテーション優先', 'どちらでも'],
  '指定タイプ': ['', '必須', '優先'],
  '割当量': ['均等', '多め', '少なめ'],
  '操作': ['キャンセル', '時間���更', '追加'],
  '制限タイプ': ['休み', '午前休', '午後休', '遅刻', '早退', '時間指定'],
  '時間指定方法': ['時間指定', '午前', '午後', '終日'],
  '曜日優先度': ['高', '中', '低'],
  '訪問頻度': ['毎週', '選択制']
};

// ============================================================
// 監査ルール定義
// ============================================================

var INPUT_AUDIT_RULES = [

  // ────────────────────────────────────────
  // 患者���スタ — ERROR（自動算出不可）
  // ────────────────────────────────────────
  { id: 'P-E01', sheet: '患者マスタ', field: 'patient_id', level: 'NG',
    message: '患者IDが未設定です', condition: '空' },
  { id: 'P-E02', sheet: '患者マスタ', field: '患者名', level: 'NG',
    message: '患者名が未入力です', condition: '空' },
  { id: 'P-E03', sheet: '患者マスタ', field: '稼働状況', level: 'NG',
    message: '稼働状況が未設定です。算出対象か判定できません', condition: '空' },
  { id: 'P-E04', sheet: '患者マスタ', field: '週訪問回数', level: 'NG',
    message: '週訪問回数が未設定です。訪問リクエストを生成できません', condition: '空または0' },
  { id: 'P-E05', sheet: '患者マスタ', field: '希望曜日', level: 'NG',
    message: '希望曜日が未選択です。訪問日を決定できません', condition: '空' },
  { id: 'P-E06', sheet: '患者マスタ', field: 'サービス時間', level: 'NG',
    message: 'サービス時間が未入力です。時間枠を計算できません', condition: '空' },
  { id: 'P-E07', sheet: '患者マスタ', field: '希���時間帯（開始）', level: 'NG',
    message: '時間タイプが「固定」ですが開始時刻が未入力です', condition: '時間タイプ=固定 かつ 開始が空' },
  { id: 'P-E08', sheet: '患者マスタ', field: '希望時間帯（開始）', level: 'NG',
    message: '時間タイプが「時間帯」ですが��間枠が未入力です', condition: '時間タイプ=時間帯 かつ 開始が空' },
  { id: 'P-E09', sheet: '患者マスタ', field: '希望時間帯（開始）', level: 'NG',
    message: '開始時刻が終了時刻以降になっています', condition: '開始 >= 終了' },
  { id: 'P-E10', sheet: '患者マスタ', field: '希望曜日', level: 'NG',
    message: '希望曜日数が週訪問回数より少ないです', condition: '曜日数 < 週訪問回数' },
  { id: 'P-E11', sheet: '患者マスタ', field: '稼働状況', level: 'NG',
    message: '稼働状況の値が不正です（稼働/休止/未開始/未契約/解約 以外）', condition: '選択肢外の値' },
  { id: 'P-E12', sheet: '患者マスタ', field: '時間タイプ', level: 'NG',
    message: '時間タイプの値が不正です（時間帯/午前/午後/終日/固定 以外）', condition: '空でなく選択肢外の値' },

  // ────────────────────────────────────────
  // 患者マスタ — WARN（精度低下）
  // ────────────────────────────────────────
  { id: 'P-W01', sheet: '患者マスタ', field: '緯度', level: 'WARN',
    message: '住所（位置情報）が未設定です。ルート最適化・距離計算に影響します', condition: '緯度または経度が空' },
  { id: 'P-W02', sheet: '患者マスタ', field: '時間タイプ', level: 'WARN',
    message: '時間タイプが未設定です。終日扱いとなり精度が低下します', condition: '空' },
  { id: 'P-W03', sheet: '患者マスタ', field: '継続希望', level: 'WARN',
    message: '継続希望が未設定です。「どちらでも」として扱われます', condition: '空' },
  { id: 'P-W04', sheet: '患者マスタ', field: '指定タイプ', level: 'WARN',
    message: '指定スタッフIDが設定されていますが指定タイプが未選択です。指定が無視されます', condition: '指定スタッフあり かつ 指定タイプ空' },
  { id: 'P-W05', sheet: '患者マスタ', field: '必要スタッフ数', level: 'WARN',
    message: '必要スタッフ数が未設定です。デフォルト（1名）で処理されます', condition: '空' },
  { id: 'P-W06', sheet: '患者マスタ', field: '性別制限', level: 'WARN',
    message: '性別制限の値が不正です（男性のみ/女性のみ/空 以外）', condition: '選択肢外の値' },
  { id: 'P-W07', sheet: '患者マスタ', field: '継続希望', level: 'WARN',
    message: '継続希望の値が不正です（同じ人希望/ローテーション優先/どちらでも 以外）', condition: '空でなく選択肢外の値' },

  // ──────────────────────────���─────────────
  // スタッフマスタ — ERROR
  // ────────────────────────────────────────
  { id: 'S-E01', sheet: 'スタッフマスタ', field: 'staff_id', level: 'NG',
    message: 'スタッフIDが未設定です', condition: '空' },
  { id: 'S-E02', sheet: 'スタッフマスタ', field: 'スタッフ名', level: 'NG',
    message: 'スタッフ名が未入力です', condition: '空' },
  { id: 'S-E03', sheet: 'スタッフマスタ', field: '勤務曜日', level: 'NG',
    message: '勤務曜日が未選択です。割当先として使用できません', condition: '空' },
  { id: 'S-E04', sheet: 'スタッフマスタ', field: 'シフト開始', level: 'NG',
    message: 'シフト開始時刻が未設定です', condition: '空' },
  { id: 'S-E05', sheet: 'スタッフマスタ', field: 'シフト終了', level: 'NG',
    message: 'シフト終了時刻が未設定です', condition: '空' },
  { id: 'S-E06', sheet: 'スタッフマスタ', field: 'シフト開始', level: 'NG',
    message: 'シフト開始が終了以降になっています', condition: '開始 >= 終了' },

  // ────────────────────────────────────────
  // スタッフマスタ — WARN
  // ────────────────────────────────────────
  { id: 'S-W01', sheet: 'スタッフマスタ', field: '性別', level: 'WARN',
    message: '性別が未設定です。性別制限のある患者に割当できません', condition: '空' },
  { id: 'S-W02', sheet: 'スタッフマスタ', field: '緯度', level: 'WARN',
    message: '拠点住所（位置情報）が未設定です。距離計算に影響します', condition: '緯度が空' },
  { id: 'S-W03', sheet: 'スタッフマスタ', field: '最大訪問件数/日', level: 'WARN',
    message: '最大訪問件数が未設定です。デフォルト（999=無制限）で処理されます', condition: '空' },
  { id: 'S-W04', sheet: 'スタッフマスタ', field: '割当量', level: 'WARN',
    message: '割当量の値が不正です（均等/多め/少なめ 以外）', condition: '空でなく選択肢外の値' },
  { id: 'S-W05', sheet: 'スタッフマスタ', field: '性別', level: 'WARN',
    message: '性別の値が不正です（男性/女性 以外）', condition: '空でなく選択肢外の値' },

  // ──────────────────────────���─────────────
  // 個別変更リクエスト — ERROR
  // ────────────────────────────────────────
  { id: 'C-E01', sheet: '個別変更リクエスト', field: 'patient_id', level: 'NG',
    message: '患者IDが未設定です', condition: '空' },
  { id: 'C-E02', sheet: '個別変更リクエスト', field: '日付', level: 'NG',
    message: '日付が未入力です', condition: '空' },
  { id: 'C-E03', sheet: '個別変更リクエスト', field: '操作', level: 'NG',
    message: '操作が未選択です', condition: '空' },
  { id: 'C-E04', sheet: '個別変更リクエスト', field: '時間タイプ', level: 'NG',
    message: '操作が「時間変更」または「追加」ですが時間タイプが未設定です', condition: '操作=時間変更/追加 かつ 時間タイプ空' },
  { id: 'C-E05', sheet: '個別変更リクエスト', field: '開始時刻(固定)', level: 'NG',
    message: '時間タイプが「固定」ですが開始/終了時刻が未入力です', condition: '時間タイプ=固定 かつ 時刻空' },
  { id: 'C-E06', sheet: '個別変更リクエスト', field: '操作', level: 'NG',
    message: '操作の値が不正です（キャンセル/時間変更/追加 以外）', condition: '空でなく選択肢外の値' },

  // ────────────────────────────────────────
  // 個別変更リクエスト — WARN
  // ────────────────────────────────────────
  { id: 'C-W01', sheet: '個別変更リクエスト', field: '希望最早', level: 'WARN',
    message: '時間タイプが「時間帯」ですが希望最早/最遅が未入力です', condition: '時間タイプ=時間帯 かつ 希望最早/最遅空' },

  // ────────────────────────────────────────
  // スタッフ個別変更リクエスト — ERROR
  // ────────────────────────────────────────
  { id: 'SC-E01', sheet: 'スタッフ個別変更リクエスト', field: 'staff_id', level: 'NG',
    message: 'スタッフIDが未設定です', condition: '空' },
  { id: 'SC-E02', sheet: 'スタッフ個別変更リクエスト', field: '日付', level: 'NG',
    message: '日付が未入力です', condition: '空' },
  { id: 'SC-E03', sheet: 'スタッフ個別変更リクエスト', field: '制限タイプ', level: 'NG',
    message: '制限タイプが未選択です', condition: '空' },
  { id: 'SC-E04', sheet: 'スタッフ個別変更リクエスト', field: '開始時刻', level: 'NG',
    message: '制限タイプが「時間指定」ですが開始/終了時刻が未入力です', condition: '時間指定 かつ 時刻空' },
  { id: 'SC-E05', sheet: 'スタッフ個別変更リクエスト', field: '開始時刻', level: 'NG',
    message: '制限タイプが「遅刻」ですが開始時刻が未入力です', condition: '遅刻 かつ 開始空' },
  { id: 'SC-E06', sheet: 'スタッフ個別変更リクエスト', field: '終了時刻', level: 'NG',
    message: '制限タイプが「早退」ですが終了時刻が未入力です', condition: '早退 かつ 終了空' },
  { id: 'SC-E07', sheet: 'スタッフ個別変更リクエスト', field: '制限タイプ', level: 'NG',
    message: '制限タイプの値が不正です', condition: '空でなく選択肢外の値' },

  // ────────────────────────────────────────
  // イベントリクエスト — ERROR
  // ────────────────────────────────────────
  { id: 'EV-E01', sheet: 'イベントリクエスト', field: 'staff_id', level: 'NG',
    message: 'スタッフIDが未設定です', condition: '空' },
  { id: 'EV-E02', sheet: 'イベントリクエスト', field: '日付', level: 'NG',
    message: '日付が未入力です', condition: '空' },
  { id: 'EV-E03', sheet: 'イベントリクエスト', field: 'タイトル', level: 'NG',
    message: 'タイトルが未入力です', condition: '空' },
  { id: 'EV-E04', sheet: 'イベントリクエスト', field: '開始時刻', level: 'NG',
    message: '時間指定方法が「時間指定」ですが開始/終了時刻が未入力です', condition: '時間指定 かつ 時刻空' },
  { id: 'EV-E05', sheet: 'イベントリクエスト', field: '所要時間(分)', level: 'NG',
    message: '時間指定方法が午前/午後/終日ですが所要時間が未入力です', condition: '午前/午後/終日 かつ 所要時間空' },
  { id: 'EV-E06', sheet: 'イベントリクエスト', field: '時間指定方法', level: 'NG',
    message: '時間指定方法が未選択です', condition: '空' },

  // ────────────────────���───────────────────
  // イベントリクエスト — WARN
  // ────────────────────────────────────────
  { id: 'EV-W01', sheet: 'イベントリクエスト', field: '住所', level: 'WARN',
    message: '住所が未設定です。移動時間の計算に影響します', condition: '空' },
  { id: 'EV-W02', sheet: 'イベントリクエスト', field: 'イベント種別', level: 'WARN',
    message: 'イベント種別が未選択です', condition: '空' },

  // ────────────────────────────────────────
  // スタッフ同行割付 — ERROR
  // ────────────────────────────────────────
  { id: 'M-E01', sheet: 'スタッフ同行割付', field: 'trainee_staff_id', level: 'NG',
    message: '研修スタッフIDが未設定です', condition: '空' },
  { id: 'M-E02', sheet: 'スタッフ同行割付', field: 'mentor_staff_id', level: 'NG',
    message: 'メンタースタッフIDが未設定です', condition: '空' },
  { id: 'M-E03', sheet: 'スタッフ同行割付', field: '開始日', level: 'NG',
    message: '開始日が未入力です', condition: '空' },
  { id: 'M-E04', sheet: 'スタッフ同行割付', field: '終了日', level: 'NG',
    message: '終了日が未入力です', condition: '空' },
  { id: 'M-E05', sheet: 'スタッフ同行割付', field: '開始日', level: 'NG',
    message: '開始日が終了日より後になっています', condition: '開始 > 終了' },
  { id: 'M-E06', sheet: 'スタッフ同行割付', field: 'trainee_staff_id', level: 'NG',
    message: '研修スタッフとメンターが同一人物です', condition: 'trainee = mentor' }
];

// ============================================================
// ヘルパー関数
// ============================================================

/**
 * シート名でルールをフィルタ
 * @param {string} sheetName
 * @return {Object[]}
 */
function inputAudit_getSheetRules(sheetName) {
  var result = [];
  for (var i = 0; i < INPUT_AUDIT_RULES.length; i++) {
    if (INPUT_AUDIT_RULES[i].sheet === sheetName) {
      result.push(INPUT_AUDIT_RULES[i]);
    }
  }
  return result;
}

/**
 * シート名 + フィールド名でルールをフィルタ
 * @param {string} sheetName
 * @param {string} fieldName
 * @return {Object[]}
 */
function inputAudit_getFieldRules(sheetName, fieldName) {
  var result = [];
  for (var i = 0; i < INPUT_AUDIT_RULES.length; i++) {
    var r = INPUT_AUDIT_RULES[i];
    if (r.sheet === sheetName && r.field === fieldName) {
      result.push(r);
    }
  }
  return result;
}

/**
 * 値が正規の選択肢に含まれるかチェック
 * @param {string} value
 * @param {string} fieldName - INPUT_AUDIT_VALID_VALUES のキー
 * @return {boolean}
 */
function inputAudit_isValidChoice(value, fieldName) {
  if (!value) return true; // 空は別ルールでチェック
  var validValues = INPUT_AUDIT_VALID_VALUES[fieldName];
  if (!validValues) return true;
  var normalized = String(value).replace(/[\s\u3000]/g, '');
  for (var i = 0; i < validValues.length; i++) {
    if (validValues[i].replace(/[\s\u3000]/g, '') === normalized) return true;
  }
  return false;
}
