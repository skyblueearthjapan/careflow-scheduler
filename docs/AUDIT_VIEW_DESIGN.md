# 監査（合否判定）ビュー 設計仕様書

## 目的

既存のGAS（週間リクエスト生成_ / 割当結果を作成_ など）を変更せず、
同一スプレッドシート・同一Webアプリに「監査（合否判定）ビュー」を追加する。

週ビュー上の実績予定が、患者マスタ／患者個別変更／イベント紐づき／特別訪問週間の条件と整合しているかを
患者単位で○/△/×判定し、患者クリックで根拠（条件と差分）を表示する。

---

## 1. 実装方針（既存コードを触らない）

- 既存のロジック・シート構造は変更しない（週間リクエスト生成_、割当結果を作成_の内部は改修しない）。
- 新規で以下のGASファイル（.gs）を追加して実装する：
  - **A) AuditController.gs** … Web APIエンドポイント（doGet / doPost or google.script.run呼び出し窓口）
  - **B) AuditDataLoader.gs** … 週単位で必要データを一括ロードしてMap化する
  - **C) AuditEvaluator.gs** … 合否判定エンジン（患者×日、患者×週）
  - **D) AuditTypes.gs** … 定数・型的な構造（キー文字列、判定種別、時間タイプ定義）
  - **E) AuditUtils.gs** … 日付/時刻パース、minutes<->serial、曜日正規化等の共通関数

- Webアプリのフロントは既存HTMLにタブを追加するだけに留める（可能なら既存UIに「監査」タブ追加）。
  - ※UI改修が重い場合は、監査専用のサブ画面（監査ページ）として追加してもよい。

---

## 2. 判定対象データ（"実績"のソース）

- 判定対象は「割当結果」シートの行を"実績予定"として扱う（visit_id, 日付, patient_id, 開始/終了 等）。
  - staff_id が未割当の行も含める（未割当はNG理由として扱う）
  - イベント行（EV_）は patient_id が空の場合があるため、患者判定の"衝突要因"として別途 eventSheet を参照する

※もし「週ビュー表示が割当結果ではなく別の表示用シート」を使っている場合は、
audit側で参照シート名を切替できるように設定化（AuditConfig）する。

---

## 3. 入力データ参照元（週単位で一括ロード）

### 必須ロード

| データ | 内容 |
|--------|------|
| 患者マスタ | 患者条件（patient_id, 希望曜日, 曜日NG, 希望時間帯開始/終了, 時間タイプ, サービス時間 等） |
| 患者個別変更リクエスト | 対象週のレコード（patient_id, 日付, 操作, 新開始/新終了, 登録日時） |
| イベントリクエスト | 対象週のレコード（患者紐づきpatient_idがあるもの、日付, start/end or timeMode等） |
| 特別訪問週間 | 対象週の設定（患者×週のADD/REPLACE、日別の訪問設定） |
| 割当結果 | 対象週の行（patient_id, 日付, 開始/終了, staff_id, 時間タイプ, 希望最早/最遅, 備考 等） |

### ロード方式

- 週開始日（Mon）を受け取って weekStart～weekEnd を決定
- 各シートは `getDataRange().getValues()` を基本にし、週範囲に該当する行だけ抽出
- 抽出後は Map 化して参照を O(1) にする

### 推奨Map例

```javascript
// patientMasterMap: pid -> {timeType, startPref, endPref, prefDays[], ngDays[], svcMin, ...}
patientMasterMap.get(pid)

// changeMap: pid|yyyy/MM/dd -> {op, newStart, newEnd, note, regAt}
// ※同日複数は登録日時で最新のみ採用
changeMap.get(`${pid}|${dateStr}`)

// eventMap: pid|yyyy/MM/dd -> [ {startMin, endMin, timeMode, patientAffect, title, ...}, ... ]
eventMap.get(`${pid}|${dateStr}`)

// specialWeekMap: pid|weekStartStr -> {mode: ADD/REPLACE, dailyPlans: {...}}
specialWeekMap.get(`${pid}|${weekStartStr}`)

// actualPlanMap: pid|yyyy/MM/dd -> [ {startMin, endMin, staffId, visitId, ...}, ... ]
// ※同日複数訪問あり得るので配列
actualPlanMap.get(`${pid}|${dateStr}`)
```

---

## 4. 判定エンジン（AuditEvaluator）

### 判定出力（UIがそのまま使えるJSON）

#### getAuditWeekSummary(weekStartStr)

```javascript
{
  weekStartStr,
  weekEndStr,
  cellSummary: {
    "staffId|yyyy/MM/dd": [
      { pid, pname, status, tags[], detailKey },
      ...
    ]
  },
  patientSummary: {
    pid: { overallStatus, ngCount, warnCount, reasons[] },
    ...
  }
}
```

#### getAuditPatientDetail(pid, weekStartStr)

```javascript
{
  pid,
  pname,
  master: { ... },
  changes: [ ...week changes... ],
  events: [ ...week events... ],
  specialWeek: { ... or null },
  actual: [
    { dateStr, staffId, startMin, endMin, visitId, ... },
    ...
  ],
  dayJudgements: [
    {
      dateStr,
      status,
      checks: [
        {
          type: "day",
          ok: true/false,
          expected: "Mon",
          actual: "Tue",
          bufferMin,
          diffMin,
          reason
        },
        {
          type: "time",
          ok: true/false,
          expectedWindow: { earliest, latest },
          actual: { start, end },
          bufferMin,
          diffMin,
          reason
        },
        {
          type: "changeApplied",
          ok: true/false,
          change: { ... },
          reason
        },
        {
          type: "eventConflict",
          ok: true/false,
          conflictEvent: { ... },
          reason
        },
        {
          type: "specialWeek",
          ok: true/false,
          mode: "ADD/REPLACE",
          reason
        }
      ]
    }
  ]
}
```

### 判定ルール（初期版：説明しやすい3段階）

| status | 説明 |
|--------|------|
| OK | 期待範囲内 |
| WARN | 期待範囲外だがバッファ内 |
| NG | バッファ超過 |

- バッファ（分）は設定値 `AUDIT_TIME_BUFFER_MIN`（例：15分）

### 時間タイプごとの期待範囲（master基準）

| 時間タイプ | 期待範囲 |
|------------|----------|
| 固定 | 開始/終了が固定（開始±buffer, 終了±bufferで判定） |
| 時間帯 | 希望開始～希望終了を期待範囲とする |
| 午前 | 09:00～12:00 |
| 午後 | 13:00～17:00 |
| 終日 | 09:00～18:00 |

※患者マスタに希望開始/終了がある場合は、それを優先し、無い場合に時間タイプのデフォルトを使う

### 個別変更の優先順位

- 同一pid・同一日付の個別変更がある場合：
  - **キャンセル**：その日の"期待"はキャンセル（実績があればNG）
  - **時間変更**：期待範囲を newStart/newEnd に置換（片方欠けたらサービス時間で補完）
  - **追加**：期待訪問が1件増える（実績が無いならWARN/NGは運用で選択。初期はWARN推奨）

- 特別訪問週間（ADD/REPLACE）は個別変更と同等以上の優先度で期待を再構成する
  - **REPLACE**: 通常期待を置換
  - **ADD**: 通常期待に追加

### イベント紐づきの扱い（患者判定として）

- eventSheetで `patientLinked=true` & `patient_id`あり のイベントを取得
- そのイベント時間と実績訪問が重なれば **NG**（理由：イベント衝突）
- イベントが「患者影響=予定変更扱い」等の区分を持つ場合は、将来拡張で期待に反映

---

## 5. API設計（フロントから呼ぶ関数）

| 関数名 | 説明 |
|--------|------|
| `audit_getWeekSummary(weekStartStr)` | 週サマリ取得 |
| `audit_getPatientDetail(pid, weekStartStr)` | 患者詳細取得 |
| `audit_getConfig()` | 設定取得（bufferや参照シート名など） |
| `audit_getCellPatients(staffId, dateStr, weekStartStr)` | （任意）セルクリック用に絞り込み |

実装は `google.script.run` で呼べるサーバ関数として公開する。
（doGet/doPostのREST形式にする場合は後からでも変更可能なように内部関数を分離）

---

## 6. パフォーマンス要件

- 週サマリ取得は「シート全走査を毎回やらない」設計にする
  - CacheService（script cache）に週単位datasetを15分程度キャッシュ
  - key例: `"auditDataset|yyyyMMdd(weekStart)"`
- patientDetail は週datasetから pid を引くだけ（再走査しない）

---

## 7. 既存コードへの影響ゼロ条件

- 既存関数名・既存シート名・既存ヘッダを変更しない
- 既存の `applySpecialWeekToWeeklyRequests_` など既存内部関数は呼ばない（監査側は"読み取り専用"）
  - ※ただし判定ロジックを一致させるため、時間タイプのデフォルト窓は既存実装に合わせる

---

## 8. 納品物（コーディングエージェントのゴール）

- [ ] 監査タブ（または監査ページ）で、週ビューのセルクリック→患者一覧（○/△/×）が出る
- [ ] 患者名クリック→個別詳細で「根拠（master/change/event/special）＋差分」が出る
- [ ] 週開始日を切り替え可能
- [ ] バッファ分を設定で変更可能（AuditConfig）

---

## ファイル構成（予定）

```
/
├── UnifiedCode.gs          # 既存（変更しない）
├── UnifiedInput.html       # 既存（監査タブ追加のみ）
├── UnifiedOutput.html      # 既存（変更しない）
├── UnifiedNoAccess.html    # 既存（変更しない）
├── appsscript.json         # 既存（変更しない）
│
├── AuditController.gs      # 新規：API エンドポイント
├── AuditDataLoader.gs      # 新規：データローダー
├── AuditEvaluator.gs       # 新規：判定エンジン
├── AuditTypes.gs           # 新規：定数・型定義
└── AuditUtils.gs           # 新規：ユーティリティ関数
```
