# Expected（期待）生成：擬似コード（patient_id × date 単位）

**目的**：監査（合否判定）で使う「その日に本来あるべき訪問枠（Expected）」を、
患者マスタ／個別変更／特別訪問週間（ADD/REPLACE）から**矛盾なく**作る。

## 前提

- 既存ロジック（週間リクエスト生成・特別訪問適用）は改変しない方針
- 監査側では「期待」を再構築する（=監査用の期待DTOを作る）
- 同一 patient_id + dateStr に複数枠（最大2〜3枠）があり得る
  - 例）通常 + 追加（個別変更ADD） + 特別ADD
- time は内部 minutes（0-1440）で扱い、入出力は適宜serial/Date/"HH:mm"を吸収

---

## 0. データ構造（監査側DTO）

### ExpectedVisit（期待枠）

```ts
type ExpectedVisit = {
  expected_id: string           // 例 "EXP|P001|2026/01/15|01"
  pid: string
  dateStr: string               // "yyyy/MM/dd"
  youbi: string                 // "Mon".."Sun"
  source: "MASTER" | "CHANGE" | "SPECIAL_ADD" | "SPECIAL_REPLACE"
  op: "" | "追加" | "時間変更" | "キャンセル" | "置換"  // 表示用（sourceと併用）
  timeType: "固定" | "時間帯" | "午前" | "午後" | "終日"
  startMin: number | null       // 固定枠の開始（なければ null）
  endMin: number | null         // 固定枠の終了（なければ null）
  earliestMin: number | null    // 許容開始の最早（時間帯/午前/午後/終日）
  latestMin: number | null      // 許容終了の最遅（時間帯/午前/午後/終日）
  svcMin: number                // サービス時間（分）
  needStaff: number             // 必要スタッフ数
  constraints: object           // 指定スタッフ/NG/性別/継続希望など（表示と判定に利用）
  note: string                  // 根拠メモ（どの行から来たか等）
}
```

### DayExpected（1日ぶん）

```ts
type DayExpected = {
  key: string                   // pid|dateStr
  visits: ExpectedVisit[]       // 期待枠（0..n）
  meta: {
    hasReplace: boolean
    hasCancel: boolean
    appliedChangeId?: string
    appliedSpecialId?: string
  }
}
```

---

## 1. 入力（監査側で読み込むもの）

| 変数名 | 内容 |
|--------|------|
| `patientMasterMap[pid]` | 患者条件（希望曜日/希望時間帯/時間タイプ/サービス時間/必要人数/制約） |
| `changeReqByPidDate[pid\|dateStr]` | 個別変更（同日最新1件を採用） |
| `specialWeekByPidDate[pid\|dateStr]` | 特別訪問週間（ADD/REPLACE、同日複数枠可） |
| `weekRange` | weekStart..weekEnd（日付） |

既存の weeklyRequests（週間リクエスト）を「期待の正」として使っても良いが、
監査では "根拠表示" のために上記ソースから組み立て直せると強い

---

## 2. ユーティリティ（擬似）

### normalizeDateStr(d): "yyyy/MM/dd"

### youbiFromDate(d): "Mon".."Sun"

### parseTimeToMin(x): number|null
Date/serial/"HH:mm"対応

### makeTimeWindow(timeTypeRaw, startPref, endPref, svcMin)

**戻り値：**
```ts
{ timeType, startMin, endMin, earliestMin, latestMin }
```

**ルール（既存仕様に合わせる）：**

| timeType | earliest | latest | start | end |
|----------|----------|--------|-------|-----|
| 固定 | - | - | 確定 | 確定（svcMinから補完可） |
| 時間帯 | 希望開始 | 希望終了 | earliest | start+svcMin |
| 午前 | 9:00 | 12:00 | 9:00 | start+svcMin |
| 午後 | 13:00 | 17:00 | 13:00 | start+svcMin |
| 終日 | 9:00 | 18:00 | 9:00 | start+svcMin |

### buildExpectedFromMaster(pid, dateStr, masterInfo) => ExpectedVisit

### buildExpectedFromChangeAdd(pid, dateStr, changeRow, masterInfo) => ExpectedVisit

### buildExpectedFromSpecialDetail(pid, dateStr, detailRow, mode, masterInfo) => ExpectedVisit

※それぞれ constraints, svcMin, needStaff を masterInfo から引き継ぐ

---

## 3. 優先順位（固定）

同一 pid|dateStr の期待生成における"適用順"：

1. **SPECIAL_REPLACE（置換）** → その日の通常期待は捨てる（REPLACEが基準になる）
2. **CHANGE（キャンセル/時間変更）** → 当日の期待枠に作用（最新1件）
3. **MASTER（通常）** → 基本期待（REPLACEが無い場合のみ）
4. **SPECIAL_ADD（追加）** → 当日に追加で積む
5. **CHANGE（追加）** → 追加で積む（※運用により SPECIAL_ADD より先でもOKだが、ここでは最後に積む）

### 注意

- CHANGE の「キャンセル」「時間変更」は「通常枠（MASTER由来）」に作用する想定が基本。
- REPLACE がある日に CHANGE がある場合：
  - 原則：REPLACE 後の期待（置換枠）に CHANGE を適用する（同日最新1件を優先）
  - ただし CHANGE が「追加」なら、置換枠に追加で積む扱いが自然

---

## 4. 擬似コード本体

### 4-1. 週の対象日を列挙

```javascript
for each date in [weekStart..weekEnd]:
  dateStr = normalizeDateStr(date)
```

### 4-2. 対象患者を列挙（方法はどちらでも）

- **A案**：患者マスタ全員を対象にする
- **B案**：割当結果（actual）や週間リクエストに出現した患者のみ対象にする（高速）

→ Phase2以降はB案推奨（無駄に重くしない）

```javascript
patients = getTargetPatients()
```

### 4-3. 期待Mapを初期化

```javascript
expectedByPidDate = Map<string, DayExpected>()  // key = pid|dateStr
```

### 4-4. 生成ループ

```javascript
for pid in patients:
  master = patientMasterMap[pid]
  if !master: continue

  for date in weekDays:
    dateStr = normalizeDateStr(date)
    key = pid + "|" + dateStr
    youbi = youbiFromDate(date)

    day = { key, visits: [], meta: {hasReplace:false, hasCancel:false} }

    // --- (1) SPECIAL_REPLACE を先に適用（あれば通常を作らない）
    specials = specialWeekByPidDate[key]  // 0..n detail rows
    replaceDetails = specials.filter(x => x.mode == "REPLACE")
    addDetails     = specials.filter(x => x.mode == "ADD")

    if replaceDetails.length > 0:
      day.meta.hasReplace = true

      // REPLACE は「明細がその日の期待」になる
      for det in replaceDetails:
        exp = buildExpectedFromSpecialDetail(pid, dateStr, det, "SPECIAL_REPLACE", master)
        day.visits.push(exp)

    else:
      // --- (2) MASTER（通常期待）を作る（その曜日が対象か）
      if isMasterVisitDay(master, youbi):  // 希望曜日 - NG曜日 判定
        exp = buildExpectedFromMaster(pid, dateStr, master)
        day.visits.push(exp)
      else:
        // その日は通常期待なし
        pass

    // --- (3) CHANGE（同日最新1件）を適用：キャンセル/時間変更/追加
    ch = changeReqByPidDate[key]  // 最新1件 or null
    if ch exists:
      if ch.op == "キャンセル":
        // 当日すでにある期待（MASTER由来でもREPLACE由来でも）をキャンセル扱いにする
        for exp in day.visits:
          exp.op = "キャンセル"
          exp.source = exp.source  // sourceは維持しつつ op でキャンセル表現
          exp.note += " | change:キャンセル"
        day.meta.hasCancel = true

      else if ch.op == "時間変更":
        // 当日期待の「メイン枠」に時間変更を当てる
        // ルール：最初の1枠だけに適用（複数枠ある場合、どれを変えるか仕様が必要）
        // ここでは「最もMASTERに近い枠」優先 → 無ければ先頭
        targetIdx = findPrimaryExpectedIndex(day.visits)
        if targetIdx >= 0:
          exp = day.visits[targetIdx]
          // 新開始/新終了があれば採用、なければsvcMinから補完
          win = makeTimeWindow(exp.timeType, ch.newStart, ch.newEnd, exp.svcMin)
          exp.startMin = win.startMin
          exp.endMin   = win.endMin
          exp.earliestMin = win.earliestMin
          exp.latestMin   = win.latestMin
          exp.op = "時間変更"
          exp.note += " | change:時間変更"
        else:
          // 期待が無いのに時間変更が来たケース → 追加扱いで積む/またはWARN
          exp = buildExpectedFromChangeAdd(pid, dateStr, ch, master)
          exp.op = "時間変更(期待無し)"
          day.visits.push(exp)

      else if ch.op == "追加":
        exp = buildExpectedFromChangeAdd(pid, dateStr, ch, master)
        exp.source = "CHANGE"
        exp.op = "追加"
        day.visits.push(exp)

    // --- (4) SPECIAL_ADD を最後に積む（REPLACE の有無に関わらず追加）
    if addDetails.length > 0:
      for det in addDetails:
        exp = buildExpectedFromSpecialDetail(pid, dateStr, det, "SPECIAL_ADD", master)
        day.visits.push(exp)

    // --- (5) 整理：キャンセルは「期待枠として残す」か「0件にする」か運用で決める
    // 監査では「キャンセルなのに実績がある」を見たいので、枠として残す推奨
    // ただし day.visits が空で、かつ CHANGE 追加も無いなら key を作らない（mapに入れない）でもOK

    if day.visits.length > 0:
      // expected_id 付番（安定化：開始時刻順で連番）
      sort day.visits by (startMin or earliestMin or 99999)
      for i in 0..len(day.visits)-1:
        day.visits[i].expected_id = "EXP|" + pid + "|" + dateStr + "|" + pad2(i+1)

      expectedByPidDate[key] = day
```

---

## 5. "メイン枠"選択（時間変更対象の選び方）

```javascript
function findPrimaryExpectedIndex(visits):
  if visits.length == 0: return -1
  // 優先：MASTER由来 > SPECIAL_REPLACE由来 > その他
  for i in range(visits):
    if visits[i].source == "MASTER": return i
  for i in range(visits):
    if visits[i].source == "SPECIAL_REPLACE": return i
  return 0
```

---

## 6. 注意点（バグの温床ポイント）

| ポイント | 説明 |
|----------|------|
| 日付正規化 | 文字列日付（"2026/1/5" 等）を必ず Date にしてから dateStr化（週範囲比較は dateStr でなく Date でも良い） |
| serial時刻 | 0〜1 は minutes に直す |
| 同日複数変更 | 同日複数の個別変更がある場合、登録日時で "最新1件のみ" にする（既存方針踏襲） |
| REPLACE二重 | REPLACE がある日に「通常」も残してしまうと二重期待になる（必ず捨てる） |
| キャンセル保持 | CHANGE の「キャンセル」を適用したら、期待枠は残して `op=キャンセル` にする（監査上重要） |
| 時間変更空振り | CHANGE の「時間変更」が、期待0件日に入っているケースの扱い（仕様決め：追加扱いにするかWARNにするか） |

---

## 7. 出力（監査UIでの使い方）

- **週ビュー（セル/患者一覧）**では、expectedByPidDate から該当日の期待を取り出し、
  actual（割当結果）と突合して ○/△/× を出す
- **患者詳細**では day.visits を「期待」として表示し、source/op/note を根拠として提示する
