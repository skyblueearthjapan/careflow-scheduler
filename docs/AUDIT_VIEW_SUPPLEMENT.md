# 監査（合否判定）ビュー 追補資料

（データキー定義 / JSONサンプル / UIワイヤー）

---

## 0. 用語（この設計書内での呼び方）

- **期待（Expected）**：患者マスタ＋個別変更＋特別訪問週間 から構築される「本来こうあるべき」予定
- **実績（Actual）**：割当結果シートにある「実際に割り当てられた」予定
- **監査（Audit）**：期待と実績の差分を○/△/×で判定し、根拠を出す

### ステータス定義

| ステータス | 説明 |
|------------|------|
| OK | 条件内 |
| WARN | 条件外だがバッファ内（±15分など） |
| NG | バッファ超過 / ルール違反 / 未割当 / 衝突 |

---

## 1. データキー定義（必須）

### 1.1 日付文字列（dateStr）

- 形式：`yyyy/MM/dd`
- 例：`2026/01/12`
- 生成：`Utilities.formatDate(dateObj, tz, 'yyyy/MM/dd')`

### 1.2 週開始（weekStartStr）

- 形式：`yyyy/MM/dd`（月曜）
- 週範囲：weekStart〜weekStart+6日（weekEndStr）
- UI/URL/キャッシュキーの中心値

### 1.3 Mapキー（共通）

#### Patient-dayキー

- `pdKey = pid + '|' + dateStr`
- 例：`P001|2026/01/12`

**用途：**
- changeMap（個別変更）
- eventMap（患者紐づきイベント）
- actualPlanMap（割当結果の患者予定）
- expectedPlanMap（期待予定）

#### Staff-dayキー（セル単位）

- `sdKey = staffId + '|' + dateStr`
- 例：`S003|2026/01/12`

**用途：**
- 週ビューのセル表示（スタッフ×日→患者一覧）

#### Patient-weekキー（特別訪問週間など）

- `pwKey = pid + '|' + weekStartStr`
- 例：`P001|2026/01/12`

**用途：**
- specialWeekMap（ADD/REPLACEのヘッダ）
- 週単位集計（patientSummary）

---

## 2. 内部データ構造（推奨スキーマ）

### 2.1 PatientMaster（患者マスタ由来）

```ts
type PatientMaster = {
  pid: string
  name: string
  area?: string
  svcMin?: number
  needStaff?: number
  sexLimit?: "" | "男性のみ" | "女性のみ"
  contPref?: "" | "どちらでも" | "同じ人希望" | "ローテーション優先"

  // 希望
  prefDays?: string[]        // ["Mon","Wed","Fri"]
  ngDays?: string[]          // ["Sat"]

  // 時間
  timeType?: "" | "固定" | "時間帯" | "午前" | "午後" | "終日"
  startPref?: any            // serial or "HH:mm" etc
  endPref?: any
}
```

### 2.2 ChangeRequest（個別変更リクエスト）

```ts
type ChangeRequest = {
  pid: string
  dateStr: string
  op: "キャンセル" | "時間変更" | "追加"
  newStart?: any
  newEnd?: any
  note?: string
  regAt?: number // ソート用（epoch ms） or rowIndex
}
```

### 2.3 SpecialWeek（特別訪問週間）

```ts
type SpecialWeek = {
  pid: string
  weekStartStr: string
  mode: "ADD" | "REPLACE"
  // 日別訪問（最大2件想定：特別1/特別2）
  dailyPlans: {
    [dateStr: string]: Array<{
      earliest?: any   // "13:00" 等
      latest?: any
      start?: any      // 固定の場合
      end?: any
      svcMin?: number
      timeType?: "固定" | "時間帯" | "午前" | "午後" | "終日"
      note?: string
    }>
  }
}
```

### 2.4 Event（イベントリクエスト：患者紐づきのみ）

```ts
type LinkedEvent = {
  eventId?: string
  pid: string
  dateStr: string
  startMin?: number
  endMin?: number
  timeMode?: "" | "午前" | "午後" | "終日" | "時間指定"
  title?: string
  patientAffect?: string  // 将来拡張
}
```

### 2.5 ActualPlan（割当結果：実績）

```ts
type ActualPlan = {
  visitId: string
  pid: string
  pname?: string
  staffId?: string
  staffName?: string
  dateStr: string
  startMin?: number
  endMin?: number
  timeType?: string
  earliestMin?: number
  latestMin?: number
  note?: string
  isUnassigned?: boolean
}
```

### 2.6 ExpectedPlan（監査側で構築する期待）

```ts
type ExpectedPlan = {
  source: "MASTER" | "CHANGE" | "SPECIAL_ADD" | "SPECIAL_REPLACE"
  pid: string
  dateStr: string

  // 期待枠（判断は minutes で行う）
  expectedStartMin?: number   // 固定的な開始（固定の場合）
  expectedEndMin?: number
  expectedEarliestMin?: number // 時間帯の下限
  expectedLatestMin?: number   // 時間帯の上限

  timeType: "固定" | "時間帯" | "午前" | "午後" | "終日"
  svcMin?: number
  note?: string
}
```

---

## 3. 判定ルール（監査ロジックの「決め方」）

### 3.1 時間の正規化

- 監査エンジン内部の基準は **minutes（0〜1440）**
- どの入力（Date/serial/"HH:mm"）でも `parseTimeToMinutes()` で minutes に統一する

### 3.2 バッファ

- `AUDIT_TIME_BUFFER_MIN = 15`
- 判定：
  - **OK**：期待範囲内
  - **WARN**：期待範囲外だが、差分がバッファ以内
  - **NG**：バッファ超過

### 3.3 優先順位（期待をどう作るか）

同一 patient/day について期待を作る優先順位：

1. **特別訪問週間 REPLACE**（あれば通常期待を置換）
2. **個別変更**：キャンセル／時間変更（同日最新）
3. **通常**（患者マスタ）
4. **特別訪問週間 ADD**（通常に追加する期待枠）

※同日に期待枠が複数あり得る（例：通常1件＋追加1件）。

### 3.4 実績との付き合わせ（複数件あり得る）

- expected[] と actual[] を同日で突合
- **初期版（実装簡易で説明しやすい）：**
  - 期待枠ごとに「最も近い実績」を1件割り当てる（貪欲でOK）
  - 未割当（actualが足りない）→ NG or WARN（運用で選択。初期は WARN 推奨）
  - 期待がキャンセルなのに実績あり → NG（重複）

### 3.5 イベント衝突

- 患者紐づきイベントがある日：
  - 実績訪問の時間区間がイベント区間と重なれば **NG**（eventConflict）

---

## 4. JSONサンプル（UIがそのまま使える形式）

### 4.1 週サマリAPI：`audit_getWeekSummary(weekStartStr)`

**レスポンス例：**

```json
{
  "weekStartStr": "2026/01/12",
  "weekEndStr": "2026/01/18",
  "bufferMin": 15,

  "cellSummary": {
    "S001|2026/01/12": [
      {
        "pid": "P001",
        "pname": "青木花子",
        "status": "OK",
        "tags": ["MASTER"],
        "detailKey": "P001|2026/01/12"
      },
      {
        "pid": "P002",
        "pname": "佐藤太郎",
        "status": "WARN",
        "tags": ["TIME_DIFF(+10m)", "CHANGE(時間変更)"],
        "detailKey": "P002|2026/01/12"
      }
    ],
    "S003|2026/01/13": [
      {
        "pid": "P001",
        "pname": "青木花子",
        "status": "NG",
        "tags": ["EVENT_CONFLICT", "OUT_OF_WINDOW"],
        "detailKey": "P001|2026/01/13"
      }
    ]
  },

  "patientSummary": {
    "P001": {
      "pname": "青木花子",
      "overallStatus": "NG",
      "okCount": 2,
      "warnCount": 0,
      "ngCount": 1,
      "reasons": ["2026/01/13: EVENT_CONFLICT"]
    },
    "P002": {
      "pname": "佐藤太郎",
      "overallStatus": "WARN",
      "okCount": 2,
      "warnCount": 1,
      "ngCount": 0,
      "reasons": ["2026/01/12: TIME_DIFF(+10m)"]
    }
  }
}
```

### 4.2 患者詳細API：`audit_getPatientDetail(pid, weekStartStr)`

**レスポンス例：**

```json
{
  "weekStartStr": "2026/01/12",
  "pid": "P001",
  "pname": "青木花子",
  "bufferMin": 15,

  "master": {
    "prefDays": ["Mon","Wed","Fri"],
    "ngDays": [],
    "timeType": "午後",
    "startPref": "13:00",
    "endPref": "17:00",
    "svcMin": 60
  },

  "changes": [
    {
      "dateStr": "2026/01/15",
      "op": "追加",
      "newStart": "15:00",
      "newEnd": "16:00",
      "note": "体調不良で追加",
      "regAt": 1736900000000
    }
  ],

  "specialWeek": {
    "mode": "ADD",
    "dailyPlans": {
      "2026/01/12": [
        { "timeType": "時間帯", "earliest": "14:00", "latest": "16:00", "svcMin": 30, "note": "特別1" }
      ]
    }
  },

  "events": [
    {
      "dateStr": "2026/01/13",
      "eventId": "E003",
      "title": "病院受診同行",
      "startMin": 840,
      "endMin": 900
    }
  ],

  "actual": [
    { "dateStr": "2026/01/12", "visitId": "V010", "staffId": "S001", "startMin": 810, "endMin": 870 },
    { "dateStr": "2026/01/13", "visitId": "V020", "staffId": "S003", "startMin": 835, "endMin": 895 }
  ],

  "dayJudgements": [
    {
      "dateStr": "2026/01/12",
      "status": "OK",
      "checks": [
        {
          "type": "timeWindow",
          "status": "OK",
          "expected": { "earliestMin": 780, "latestMin": 1020 },
          "actual": { "startMin": 810, "endMin": 870 },
          "diffMin": 0,
          "reason": ""
        },
        {
          "type": "specialWeek",
          "status": "OK",
          "reason": "SPECIAL_ADD exists (additional expectation created)"
        }
      ]
    },
    {
      "dateStr": "2026/01/13",
      "status": "NG",
      "checks": [
        {
          "type": "eventConflict",
          "status": "NG",
          "event": { "startMin": 840, "endMin": 900, "title": "病院受診同行" },
          "actual": { "startMin": 835, "endMin": 895 },
          "reason": "overlaps linked event"
        }
      ]
    }
  ]
}
```

---

## 5. UIワイヤー（テキスト / 迷わない構造）

### 5.1 画面追加方針

- 既存の「入力/出力」等のナビは維持
- 追加：**「監査」**タブ（またはサブページ）
- 監査は "読み取り専用" を基本（編集はしない）

### 5.2 監査 週ビュー（メイン）

```
┌───────────────────────────────────────────────┐
│ [監査] 週開始: (DatePicker) [2026/01/12] [更新] │
│ buffer: 15分   凡例: ○OK △WARN ×NG             │
├───────────────────────────────────────────────┤
│ 週ビュー（スタッフ×日）                          │
│     Mon        Tue        Wed ...              │
│ S001 [○2 △1 ×0] [○1 △0 ×1] ...                 │
│ S002 ...                                        │
│ S003 ...                                        │
├───────────────────────────────────────────────┤
│ セル詳細（クリックで右 or 下に展開）              │
│ 「S001 / 2026/01/12 の患者」                      │
│  - ○ 青木花子  [MASTER]          (詳細ボタン)     │
│  - △ 佐藤太郎  [CHANGE][+10m]     (詳細ボタン)     │
│  - × 田中一郎  [OUT_OF_WINDOW]    (詳細ボタン)     │
└───────────────────────────────────────────────┘
```

**UI要件：**
- セルには「○△×件数」を表示（一覧性）
- セルクリックで患者一覧（status + tags）
- 患者クリックで患者詳細モーダル or 右ペイン

### 5.3 患者詳細（モーダル/右ペイン）

```
┌───────────────────────────────┐
│ 患者: 青木花子 (P001) 週: 2026/01/12            │
├───────────────────────────────┤
│ [患者マスタ]                                      │
│  希望曜日: Mon/Wed/Fri  時間タイプ: 午後          │
│  希望: 13:00-17:00  サービス:60                   │
├───────────────────────────────┤
│ [個別変更]                                        │
│  2026/01/15 追加 15:00-16:00  備考...             │
├───────────────────────────────┤
│ [特別訪問週間]                                    │
│  ADD: 2026/01/12 14:00-16:00 (30m)                │
├───────────────────────────────┤
│ [イベント(患者紐づき)]                             │
│  2026/01/13 14:00-15:00 病院受診同行              │
├───────────────────────────────┤
│ [合否判定（1週間）]                                │
│  01/12 ○ 期待:13-17 実績:13:30-14:30 差分0         │
│  01/13 × イベント衝突（14:00-15:00 と重複）         │
│  01/15 △ 追加期待あり／実績不足（運用でWARN）      │
└───────────────────────────────┘
```

**表示ルール：**
- "根拠データ" をブロック別に表示（master/change/special/event）
- "判定" は日単位に並べる（見落とし防止）
- tags（原因ラベル）を短い文字列で出す
  - `TIME_DIFF(+10m)`
  - `OUT_OF_WINDOW`
  - `CHANGE(キャンセル/時間変更/追加)`
  - `SPECIAL(ADD/REPLACE)`
  - `EVENT_CONFLICT`
  - `UNASSIGNED`

---

## 6. 監査タグ（tags）標準案

### OK系
- `MASTER`
- `CHANGE(時間変更)`
- `SPECIAL_ADD` / `SPECIAL_REPLACE`

### WARN系
- `TIME_DIFF(+Xm)`
- `MISSING_ACTUAL`（期待があるのに実績が無い）
- `EXTRA_ACTUAL`（期待が無いのに実績がある）

### NG系
- `UNASSIGNED`
- `OUT_OF_WINDOW`
- `EVENT_CONFLICT`
- `CANCELLED_BUT_VISITED`（キャンセル期待なのに実績あり）

---

## 7. 実装補助（キャッシュキー）

- `datasetCacheKey = auditDataset|${weekStartStr.replaceAll('/', '')}`
- 例：`auditDataset|20260112`
- cache TTL：15分推奨

---

## 8. 例：セルクリック→患者一覧の最小DTO

（UIはこれだけでまず動く）

```json
{
  "sdKey": "S001|2026/01/12",
  "items": [
    {"pid":"P001","pname":"青木花子","status":"OK","tags":["MASTER"]},
    {"pid":"P002","pname":"佐藤太郎","status":"WARN","tags":["CHANGE(時間変更)","TIME_DIFF(+10m)"]}
  ]
}
```

---

## 9. 例：患者クリック→詳細の最小DTO

```json
{
  "pid":"P001",
  "pname":"青木花子",
  "weekStartStr":"2026/01/12",
  "master":{"prefDays":["Mon","Wed","Fri"],"timeType":"午後","startPref":"13:00","endPref":"17:00"},
  "dayJudgements":[
    {"dateStr":"2026/01/12","status":"OK","tags":["MASTER"]},
    {"dateStr":"2026/01/13","status":"NG","tags":["EVENT_CONFLICT"]}
  ]
}
```

---

## まとめ

この追補を参照することで、エージェントは以下を迷わず実装できます：

- **キー設計**（pid|date、staff|date、pid|week）
- **UIが要求するJSON形**
- **画面の遷移と表示要件**
