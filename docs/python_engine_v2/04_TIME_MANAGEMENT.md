# 04. 時間スロット管理・GapPack・ルート最適化設計

## TimeSlotManager

SortedList[Interval]基盤の統一的な時間管理。

### Interval

```python
@dataclass(frozen=True, order=True)
class Interval:
    start: int              # 開始分
    end: int                # 終了分
    visit_id: str
    patient_id: str
    kind: IntervalKind      # EVENT / FIXED / COUPLED / FLEX
    coupled_pair_id: Optional[str]
```

### StaffDayTimeline

スタッフ1名×1日の時間軸管理。O(log n) insert/remove。

```python
class StaffDayTimeline:
    insert(interval) -> bool
    remove(visit_id) -> Optional[Interval]
    query_gaps(min_duration, time_window) -> List[Gap]
    find_earliest_slot(duration, time_window, patient_id) -> Optional[int]
    has_overlap(start, end, exclude_visit_id) -> bool
```

### TimeSlotManager（統合管理）

```python
class TimeSlotManager:
    get_timeline(staff_id, date_str) -> StaffDayTimeline
    register_visit(staff_id, date_str, visit_id, ...) -> bool
    unregister_visit(staff_id, date_str, visit_id) -> Optional[Interval]
    find_all_overlaps(staff_id, date_str) -> List[Tuple[Interval, Interval]]
```

## BufferPolicy

距離依存バッファ（二重適用を排除）。

| 距離 | バッファ |
|------|---------|
| 同一患者 | 0分 |
| 2km以下 | 5分 |
| 2-10km | 10分 |
| 10km超 | 15分 |
| 距離不明 | 10分 |

バッファは訪問ペア間で**1回だけ**適用（GAS版の二重適用を修正）。

## TimeWindow

```python
class TimeWindow:
    earliest: int
    latest: int

    @staticmethod
    from_time_type(time_type, explicit_earliest, explicit_latest) -> TimeWindow
    # 午前: 540-720, 午後: 780-1020, 終日: 540-1080
```

## 改善版GapPack (GapPacker)

### GAS版からの改善点
1. バッファはBufferPolicy経由で1回だけ適用
2. 可動訪問配置時にTimeWindowを必ずチェック
3. 配置失敗時はunregister_visit()で確実に除去
4. 2名体制ペアは同時配置

### 分類ルール
- EVENT → アンカー
- is_fixed + 時刻確定 → アンカー
- coupled_pair_id あり → coupled_groups（ペア単位管理）
- それ以外 → 可動（flex）

### 2名体制ペア同時配置
1. ペアの共通時間窓を計算
2. 両スタッフのタイムラインから共通空きギャップを抽出
3. 最早の共通スロットにペア全体を配置
4. 片方でも不可なら両方unregister

## 改善版ルート最適化 (RouteOptimizer)

### GAS版からの改善点
1. 時間窓付き訪問も最適化対象
2. 2-opt swap後に時間窓チェック→違反なら棄却
3. 距離依存バッファ
4. 2名体制もペア単位でルートに参加

### アルゴリズム
1. 時間窓グループ分割（午前/午後/終日）
2. グループ内でnearest_neighbor + 2-opt
3. 2-opt swap後にsimulate_times()で時間窓制約チェック
4. reassign_times()でTimeSlotManagerに書き戻し

### 時間窓付き2-opt疑似コード
```
for i in 1..len-2:
  for k in i+1..len-1:
    candidate = reverse(route[i..k])
    if route_distance(candidate) < route_distance(best):
      times = simulate_times(candidate)
      if all(tw.contains(t.start, t.end) for t, tw in times):
        best = candidate  # 時間窓OK → 採用
```

## 改善版OverlapFix (OverlapResolver)

### GAS版からの改善点
1. 2名体制も含めた統合的重複解消
2. シフト方向選択（前詰め/後詰め/最短）
3. ペア全体を同時にずらす

### 優先度
EVENT > COUPLED > FIXED > FLEX

低優先度側をずらすか未割当にする。2名体制ならペア全体を同時にずらす。

## 計算量比較

| 操作 | GAS版 | Python版 |
|------|-------|---------|
| 区間挿入 | O(1) push + O(n) 重複チェック | O(log n) SortedList.add |
| 区間除去 | 未実装（忘れの原因） | O(log n) SortedList.remove |
| 空きギャップクエリ | O(n) 毎回再計算 | O(n) 構造化イテレーション |
| 重複検出 | O(n^2) 全ペアチェック | O(n log n) ソート済みスキャン |
