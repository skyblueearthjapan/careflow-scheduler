# 監査ビュー実装：シート名・列名マッピング

コードベース（UnifiedCode.gs）から抽出した実際のシート名と列名の定義。
監査機能の実装時はこのドキュメントを参照すること。

---

## 1. シート名定義（SHEETS定数）

```javascript
const SHEETS = {
  // 出力系
  WEEK_VIEW: '週ビュー',
  WEEKLY_REQUEST: '週間リクエスト',
  ASSIGN_RESULT: '割当結果',
  ASSIGN_NG: '割当不可',
  ROUTE_SUMMARY: 'ルートサマリ',
  // 入力系
  PATIENT_MASTER: '患者マスタ',
  STAFF_MASTER: 'スタッフマスタ',
  CHANGE_REQUEST: '個別変更リクエスト',
  STAFF_CHANGE_REQUEST: 'スタッフ個別変更リクエスト',
  EVENT_REQUEST: 'イベントリクエスト',
  // 権限
  ADMIN: '管理者',
  // その他
  LOG: '実行ログ'
};

// 特別訪問週間（別定義）
const SPECIAL_WEEK_HEADER = '特別訪問週間_ヘッダ';
const SPECIAL_WEEK_DETAIL = '特別訪問週間_明細';
```

---

## 2. 患者マスタ（患者マスタ）

**監査で使用する列：**

| 列名 | 用途 | 備考 |
|------|------|------|
| `patient_id` | 患者ID | 主キー |
| `患者名` | 患者名 | 表示用 |
| `エリア` | エリア | 参考情報 |
| `週訪問回数` | 週の訪問回数 | 期待生成に使用 |
| `希望曜日（複数可）` | 希望する曜日 | 複数可、カンマ区切り等 |
| `曜日NG` | 訪問不可曜日 | 複数可 |
| `時間タイプ` | 固定/時間帯/午前/午後/終日 | 期待時間帯の判定に使用 |
| `希望時間帯（開始）` | 希望開始時刻 | serial or HH:mm |
| `希望時間帯（終了）` | 希望終了時刻 | serial or HH:mm |
| `サービス時間` | サービス時間（分） | 期待終了時刻の補完に使用 |
| `必要スタッフ数` | 同行訪問用 | 通常は1 |
| `性別制限` | 男性のみ/女性のみ/空 | 表示用 |
| `継続希望` | 継続希望有無 | 表示用 |
| `指定スタッフID` | 指定スタッフ | 表示用 |
| `指定タイプ` | 必須/優先/空 | 表示用 |
| `NGスタッフID` | NGスタッフ | 表示用 |

**コード例（既存実装から）：**
```javascript
const pIdx = {
  id: pHeader.indexOf('patient_id'),
  name: pHeader.indexOf('患者名'),
  area: pHeader.indexOf('エリア'),
  svcMin: pHeader.indexOf('サービス時間'),
  timeType: pHeader.indexOf('時間タイプ'),
  startPref: pHeader.indexOf('希望時間帯（開始）'),
  endPref: pHeader.indexOf('希望時間帯（終了）'),
  prefDays: pHeader.indexOf('希望曜日（複数可）'),
  ngDays: pHeader.indexOf('曜日NG'),
  needStaff: pHeader.indexOf('必要スタッフ数'),
  sexLimit: pHeader.indexOf('性別制限'),
  contPref: pHeader.indexOf('継続希望'),
  fixedStaff: pHeader.indexOf('指定スタッフID'),
  fixedType: pHeader.indexOf('指定タイプ'),
  ngStaff: pHeader.indexOf('NGスタッフID')
};
```

---

## 3. 個別変更リクエスト（個別変更リクエスト）

**監査で使用する列：**

| 列名 | 用途 | 備考 |
|------|------|------|
| `patient_id` | 患者ID | キー |
| `患者名` | 患者名 | 表示用 |
| `日付` | 対象日付 | Date型 or 文字列 |
| `操作（キャンセル/時間変更/追加）` | 操作種別 | ヘッダー表記揺れあり（「操作」でも検索可） |
| `新開始時刻` | 新しい開始時刻 | 時間変更・追加時 |
| `新終了時刻` | 新しい終了時刻 | 時間変更・追加時 |
| `備考` | 備考 | 表示用 |
| `登録日時` | 登録日時 | 同日複数時の最新判定用 |

**操作値：**
- `キャンセル` - その日の訪問をキャンセル
- `時間変更` - 訪問時刻を変更
- `追加` - 訪問を追加

**コード例（既存実装から）：**
```javascript
const cIdx = {
  patient_id: findHeaderIdx_(cHeader, 'patient_id'),
  name:       findHeaderIdx_(cHeader, '患者名'),
  date:       findHeaderIdx_(cHeader, '日付'),
  op:         findHeaderIdx_(cHeader, '操作'),  // 部分一致で検索
  newStart:   findHeaderIdx_(cHeader, '新開始'),
  newEnd:     findHeaderIdx_(cHeader, '新終了'),
  note:       findHeaderIdx_(cHeader, '備考'),
  regAt:      findHeaderIdx_(cHeader, '登録日時')
};
```

**注意：** ヘッダー名は表記揺れがあるため、部分一致で検索する既存の `findHeaderIdx_` 関数を参考にすること。

---

## 4. イベントリクエスト（イベントリクエスト）

**監査で使用する列（患者紐づきイベント）：**

| 列名 | 用途 | 備考 |
|------|------|------|
| `event_id` | イベントID | 識別用 |
| `staff_id` | スタッフID | キー |
| `日付` | 対象日付 | Date型 or 文字列 |
| `イベント種別` | 種別 | 表示用 |
| `タイトル` | タイトル | 表示用 |
| `時間指定方法` | 午前/午後/終日/時間指定 | 時間帯判定用 |
| `開始時刻` | 開始時刻 | serial or HH:mm |
| `終了時刻` | 終了時刻 | serial or HH:mm |
| `所要時間(分)` | 所要時間 | 分単位 |
| `固定枠` | TRUE/FALSE | 固定枠かどうか |
| `患者紐づき` | TRUE/FALSE | **監査で重要：患者紐づきイベントの判定** |
| `patient_id` | 患者ID | 患者紐づき時の患者ID |
| `患者影響` | 影響種別 | 将来拡張用 |

**コード例（既存実装から）：**
```javascript
const evIdx = {
  eventId: findEvHeaderIndex(evHeader, 'event_id'),
  staffId: findEvHeaderIndex(evHeader, 'staff_id'),
  date: findEvHeaderIndex(evHeader, '日付'),
  eventType: findEvHeaderIndex(evHeader, 'イベント種別'),
  title: findEvHeaderIndex(evHeader, 'タイトル'),
  timeMode: findEvHeaderIndex(evHeader, '時間指定方法'),
  startTime: findEvHeaderIndex(evHeader, '開始時刻'),
  endTime: findEvHeaderIndex(evHeader, '終了時刻'),
  durationMin: findEvHeaderIndex(evHeader, '所要時間'),
  fixedSlot: findEvHeaderIndex(evHeader, '固定枠'),
  patientLinked: findEvHeaderIndex(evHeader, '患者紐づき'),
  patientId: findEvHeaderIndex(evHeader, 'patient_id'),
  patientAffect: findEvHeaderIndex(evHeader, '患者影響')
};
```

**患者紐づきイベントの抽出条件：**
```javascript
// 患者紐づきイベントのみ抽出
if (evRow[evIdx.patientLinked] === true || evRow[evIdx.patientLinked] === 'TRUE') {
  // patientId があれば衝突判定に使用
}
```

---

## 5. 特別訪問週間_ヘッダ

**監査で使用する列：**

| 列名 | 用途 | 備考 |
|------|------|------|
| `special_week_id` | 特別訪問週間ID | 主キー |
| `patient_id` | 患者ID | キー |
| `患者名` | 患者名 | 表示用 |
| `週開始日` | 週開始日 | Date型 |
| `週終了日` | 週終了日 | Date型 |
| `適用モード(ADD/REPLACE)` | ADD/REPLACE | **監査で重要：期待生成に影響** |
| `理由` | 理由 | 表示用 |
| `状態` | 状態 | active等 |
| `登録日時` | 登録日時 | 参考情報 |

**コード例：**
```javascript
const hIdx = {
  specialId: hHeader.indexOf('special_week_id'),
  patientId: hHeader.indexOf('patient_id'),
  patientName: hHeader.indexOf('患者名'),
  weekStart: hHeader.indexOf('週開始日'),
  weekEnd: hHeader.indexOf('週終了日'),
  mode: hHeader.indexOf('適用モード(ADD/REPLACE)'),  // 表記揺れ：'モード' の場合もあり
  reason: hHeader.indexOf('理由'),
  status: hHeader.indexOf('状態')
};
```

---

## 6. 特別訪問週間_明細

**監査で使用する列：**

| 列名 | 用途 | 備考 |
|------|------|------|
| `special_week_id` | 特別訪問週間ID | ヘッダとの紐づけ |
| `patient_id` | 患者ID | キー |
| `日付` | 対象日付 | Date型 |
| `曜日` | 曜日 | Mon〜Sun or 日本語 |
| `行ラベル` | 特別1/特別2等 | 複数枠の識別 |
| `時間タイプ` | 固定/時間帯/午前/午後/終日 | 表記揺れ：'timeType' の場合もあり |
| `開始時刻` | 開始時刻 | 固定の場合 |
| `終了時刻` | 終了時刻 | 固定の場合 |
| `希望最早` | 最早時刻 | 表記揺れ：'希望最早時刻' の場合もあり |
| `希望最遅` | 最遅時刻 | 表記揺れ：'希望最遅時刻' の場合もあり |
| `サービス時間` | サービス時間（分） | 期待終了の補完用 |
| `必要スタッフ数` | 必要スタッフ数 | 通常1 |
| `備考` | 備考 | 表示用 |
| `個別変更の扱い` | 置換する/残す | REPLACEモード時の挙動 |

**コード例（新フォーマット対応）：**
```javascript
const isNewFormat = dHeader.indexOf('special_week_id') >= 0;
const dIdx = {
  specialId: isNewFormat ? dHeader.indexOf('special_week_id') : dHeader.indexOf('special_id'),
  patientId: dHeader.indexOf('patient_id'),
  date: dHeader.indexOf('日付'),
  dayOfWeek: dHeader.indexOf('曜日'),
  rowLabel: dHeader.indexOf('行ラベル'),
  timeType: isNewFormat ? dHeader.indexOf('時間タイプ') : dHeader.indexOf('timeType'),
  start: dHeader.indexOf('開始時刻'),
  end: dHeader.indexOf('終了時刻'),
  earliest: isNewFormat ? dHeader.indexOf('希望最早') : dHeader.indexOf('希望最早時刻'),
  latest: isNewFormat ? dHeader.indexOf('希望最遅') : dHeader.indexOf('希望最遅時刻'),
  svcMin: dHeader.indexOf('サービス時間'),
  needStaff: dHeader.indexOf('必要スタッフ数'),
  note: dHeader.indexOf('備考'),
  changeHandle: dHeader.indexOf('個別変更の扱い')
};
```

---

## 7. 割当結果（割当結果）

**監査で使用する列（実績＝Actual）：**

| 列名 | 用途 | 備考 |
|------|------|------|
| `visit_id` | 訪問ID | 主キー、EV_で始まるものはイベント |
| `日付` | 訪問日付 | Date型 |
| `曜日` | 曜日 | Mon〜Sun or 日本語 |
| `staff_id` | スタッフID | キー（空＝未割当） |
| `スタッフ名` | スタッフ名 | 表示用 |
| `patient_id` | 患者ID | キー（EV行は空の場合あり） |
| `患者名` | 患者名 | 表示用 |
| `エリア` | エリア | 参考情報 |
| `開始時刻` | 開始時刻 | serial |
| `終了時刻` | 終了時刻 | serial |
| `サービス時間` | サービス時間（分） | 参考情報 |
| `時間タイプ` | 固定/時間帯/午前/午後/終日 | 判定に使用 |
| `希望最早時刻` | 最早時刻 | 期待との比較用 |
| `希望最遅時刻` | 最遅時刻 | 期待との比較用 |
| `備考` | 備考 | 表示用 |

**ヘッダー定義（既存コードから）：**
```javascript
const header = [
  'visit_id', '日付', '曜日', 'staff_id', 'スタッフ名',
  'patient_id', '患者名', 'エリア', '開始時刻', '終了時刻',
  'サービス時間', '時間タイプ', '希望最早時刻', '希望最遅時刻', '備考',
  '前訪問ID', '移動距離(km)', '移動時間(分)'
];
```

---

## 8. 時間変換ユーティリティ

**既存の時間変換関数を参考にすること：**

```javascript
// シリアル値 → 分（0〜1440）
function toMinutes(val) {
  // Date型、serial値（0〜1）、"HH:mm"文字列に対応
  // 既存実装: parseTimeToMinutes_ を参照
}

// 分 → シリアル値
function minutesToSerial_(min) {
  return min / 1440;
}

// 日付パース（表記揺れ対応）
function parseDateLoose_(val) {
  // Date型、文字列（yyyy/MM/dd, yyyy-MM-dd等）に対応
}
```

---

## 9. 時間タイプのデフォルト範囲

**既存実装に合わせること：**

| 時間タイプ | earliestMin | latestMin |
|------------|-------------|-----------|
| 午前 | 540 (9:00) | 720 (12:00) |
| 午後 | 780 (13:00) | 1020 (17:00) |
| 終日 | 540 (9:00) | 1080 (18:00) |
| 固定 | startPref | endPref |
| 時間帯 | startPref | endPref |

**コード例（既存実装から）：**
```javascript
var tt = String(row[11] || '').trim(); // 時間タイプ
if (earliestMin == null || latestMin == null) {
  if (tt === '午前') { earliestMin = 9 * 60;  latestMin = 12 * 60; }
  else if (tt === '午後') { earliestMin = 13 * 60; latestMin = 17 * 60; }
  else if (tt === '終日') { earliestMin = 9 * 60; latestMin = 18 * 60; }
}
```

---

## 10. 監査用に追加が必要なシート定義

**監査機能で新たに必要になる可能性のある定数：**

```javascript
// 監査タブ定義（OUTPUT_TABSに追加）
const AUDIT_TAB = { key: 'audit', name: '監査', sheetName: null };  // シートは不要（API経由）

// 監査設定
const AUDIT_CONFIG = {
  TIME_BUFFER_MIN: 15,  // バッファ（分）
  CACHE_TTL_SEC: 900,   // キャッシュTTL（秒）= 15分
  CACHE_KEY_PREFIX: 'auditDataset|'
};
```

---

## まとめ

監査機能の実装時は、以下に注意すること：

1. **シート名**：SHEETS定数を使用（特別訪問週間は別途定義）
2. **列名の表記揺れ**：部分一致検索を使用（findHeaderIdx_パターン）
3. **時間変換**：既存の toMinutes / parseTimeToMinutes_ を参考に
4. **日付変換**：parseDateLoose_ パターンで表記揺れに対応
5. **Boolean判定**：TRUE / 'TRUE' の両方に対応
