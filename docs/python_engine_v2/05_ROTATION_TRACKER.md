# 05. ローテーション・継続性ロジック設計

## RotationTracker クラス

過去N週の確定履歴からローテーション制御を行う。__init__でO(H)インデックス構築、スコア算出はO(1)。

### インデックス構造

```python
_profiles[pid][sid] = PatientStaffProfile
    # total_count, recent_weeks, last_date, consecutive_weeks

_recent_counts[pid] = Counter(sid -> count)   # lookback期間内
_recent_sequence[pid] = [staff_id, ...]       # 直近N回（新しい順）
_dominant_staff[pid] = staff_id               # 最頻担当
_ever_assigned[pid] = set(staff_id)           # 全期間担当経験
_current_week_assignments[pid] = [staff_id, ...]  # 週内動的追跡
```

### LOOKBACK_WEEKS = 4

## get_rotation_score()

```python
get_rotation_score(patient_id, staff_id, continuation_pref, prev_staff_id, date_str)
    -> float (-100 ~ +100)
```

### 「同じ人希望」スコアリング

| 条件 | スコア |
|------|--------|
| dominant_staff一致 | +80 + consecutive_bonus(最大+15) |
| prev_staff_id一致 | +60 |
| 担当経験あり | +20 * (担当回数/最大回数) |
| 未経験スタッフ | -30 |

### 「ローテーション優先」スコアリング

| 条件 | スコア |
|------|--------|
| 過去N週で未担当 | +60 |
| 担当少ない | +40 * (1 - 回数/最大回数) |
| 直近担当ペナルティ | -20 * recency_factor |
| 3週連続同じ人 | -50（強制ペナルティ） |
| 今週割当済み | -40（週内分散） |
| 曜日シフトボーナス | +10（ハッシュベース） |

### 「どちらでも」スコアリング

| 条件 | スコア |
|------|--------|
| 今週未割当 | +20 |
| 今週割当済み | -10 * 回数 |
| 担当経験あり | +5 |

## should_force_rotate()

3週連続で同じスタッフなら除外推奨。呼び出し側で候補から除外するか、フォールバックとして使用。

## record_assignment()

Level0処理中に呼び出し、週内の割当を動的に記録。

## スコアリング統合

```python
total = pref_score(0/1000)
      + same_today_penalty(0/-500)
      + rotation_score(-100~+100)     # RotationTracker由来
      + load_balance_score(-50~0)
      + distance_score(-10~0)
```

## 曜日シフトボーナス

GAS版の`dayOfWeek % candidates.length`はcandidates数依存で不安定。
ハッシュベース（pid+sid+曜日）で決定論的かつ安定なボーナスに改善。
