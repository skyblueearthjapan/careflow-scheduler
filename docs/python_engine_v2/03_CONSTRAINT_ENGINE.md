# 03. 制約エンジン・フィルタリング設計

## ConstraintChecker クラス

全ハード制約を統一チェック。Level0/Level1/任意のフェーズから同一メソッドで呼び出す。

### メソッド

```python
check_all(staff_id, visit, date_str, relaxed_constraints=None) -> ConstraintResult
```

### ハード制約 H1-H11

| ID | 制約名 | 内容 | GAS版の状態 |
|----|--------|------|-------------|
| H1 | NGスタッフ除外 | ng_staff_idsに含まれるスタッフは割当不可 | Level1で欠落 |
| H2 | 性別制限 | 女性のみ/男性のみ | Level1で欠落 |
| H3 | 勤務曜日 | work_daysに該当日が含まれるか | OK |
| H4 | シフト時間内 | shift_start〜shift_end内 | Level1で不完全 |
| H5 | スタッフ個別変更 | 休み/午前休/午後休/遅刻/早退 | Level1で休みのみ |
| H6 | maxPerDay | soft_cap/max_per_dayの統一 | 不整合あり |
| H7 | 時間重複禁止 | 同一スタッフ・同時間帯 | OK |
| H8 | 2名体制原子性 | 片方未割当なら両方未割当 | 事後修正のみ |
| H9 | 時間窓 | 午前/午後/終日/固定 | Level3で無視 |
| H10 | メンターペア | 同行訪問の制約 | 未実装 |
| H11 | 必須スタッフ | specified_type=必須のスタッフのみ | OK |

### 午前休/午後休の処理

```
AM_PM_BOUNDARY = 720 (12:00) - 業界慣行で固定
LUNCH_BREAK_END = 780 (13:00) - 午前休明けの実効開始

午前休: blocked = [shift_start, max(720, LUNCH_BREAK_END)]
午後休: blocked = [720, shift_end]
```

### ConstraintResult / Violation

```python
@dataclass
class Violation:
    constraint_id: str  # "H1", "H2", ...
    message: str        # 人間可読な理由
    severity: str       # "hard" | "soft"

@dataclass
class ConstraintResult:
    feasible: bool
    violations: list[Violation]
```

## ScheduleState クラス

割当済み訪問の状態を追跡。

```python
class ScheduleState:
    _assign_count: dict[str, int]          # "staff_id|date" -> count
    _staff_day_visits: dict[str, list]     # "staff_id|date" -> visits
    _patient_week_count: dict[str, int]    # "pid|staff_id" -> count
    _last_assigned_staff: dict[str, str]   # pid -> last_staff_id
    _coupled_map: dict[str, list[int]]     # base_visit_id -> [idx1, idx2]

    def add_assignment(staff_id, visit, date_str)
    def remove_assignment(staff_id, visit, date_str)
    def get_assign_count(staff_id, date_str) -> int
```

## ScoringEngine クラス

ソフト制約をスコア化し、候補スタッフをランキング。

### ソフト制約 S1-S7

| ID | 制約名 | 重み | スコア範囲 |
|----|--------|------|-----------|
| S1 | 指定スタッフ優先 | 100 | 0.0(一致) / 1.0(不一致) |
| S2 | 同じ人希望 | 80 | 0.0(前回スタッフ) / 1.0(別) |
| S3 | ローテーション | 70 | 0.0(未担当) / 1.0(連続) |
| S4 | 距離最小化 | 20 | 0.0(2km以下) / 1.0(10km超) |
| S5 | 負荷バランス | 30 | 0.0(最少) / 1.0(最多) |
| S6 | エリアマッチング | 10 | 0.0(一致) / 1.0(不一致) |

```python
rank_candidates(candidates, visit, date_str) -> list[ScoredCandidate]
```

## RelaxationPolicy

段階的緩和で割当不可を最小化。

| Level | 緩和内容 |
|-------|---------|
| 0 | 全制約適用（デフォルト） |
| 1 | S3(ローテーション)無視 |
| 2 | S2(同じ人希望)無視 |
| 3 | S6(エリア)無視 |
| 4 | H6をsoft_cap→max_per_dayに緩和 |
| 5 | S4(距離)無視 |
| 6 | S5(負荷バランス)無視 |
| 7 | H6完全無視 |
| 8 | H9(時間窓)緩和（午前/午後→終日に拡大） |

**絶対に緩和しない**: H1(NG), H2(性別), H3(勤務曜日), H5(休み), H7(時間重複), H8(ペア原子性)

## Level0/Level1統合フロー

両方とも`check_all()`を使用し、制約欠落を構造的に不可能にする。

```python
# Level0
for visit in sorted_requests:
    for staff in all_staff:
        result = constraint_checker.check_all(staff.id, visit, date_str)
        if result.feasible:
            feasible_staff.append(staff)
    ranked = scoring_engine.rank_candidates(feasible_staff, visit, date_str)

# Level1: 同一のcheck_allを使用
for visit in unassigned:
    for staff in all_staff:
        result = constraint_checker.check_all(staff.id, visit, date_str)  # 全く同じ
        if result.feasible:
            ...
```
