# 06. 2名体制・同行研修の統合設計

## CoupledVisitManager

2名体制（need_staff=2）訪問のペア原子性を全パイプライン段階で保証。

### 主要メソッド

| メソッド | 役割 |
|---------|------|
| build_pair_map() | V\d+-\d+ regexでペアマップ構築 |
| enforce_atomicity() | 片方未割当→両方未割当 |
| atomic_assign(base_vid, staff1, staff2, start, end) | 両方同時割当、片方失敗→両方棄却 |
| atomic_unassign(base_vid) | 両方同時解除 |
| sync_all_pairs() | 共通空きスロット探索→時刻同期 |
| get_anchor_indices() | GapPack/Level3でアンカー扱い |
| verify_post_gappack/level1/level3() | 各段階後の検証 |

### _find_common_free_slot()

両スタッフの空きギャップを交差して共通空き時間を探索。

```
1. 各スタッフの占有区間+バッファを収集
2. 空きギャップを計算
3. ギャップリストを交差
4. サービス時間が入る最早の共通スロットを返す
```

### select_staff_pair()（Level0用）

2名同時選択:
1. 全候補をスコア順ソート
2. Top候補(staff1)の最早空き時刻を計算
3. 残りの候補からstaff1の時刻で空いているstaff2を探す
4. ペアが見つからなければ次のstaff1候補へ

## MentorPairProcessor

割当完了後（Phase 14）にmentorの訪問をtraineeにコピー。

### GAS版からの改善点

| GAS版 | Python版 |
|-------|---------|
| traineeの独自訪問を全削除 | 独自訪問を保持 |
| 時間衝突チェックなし | 衝突チェック+flex訪問ずらし |
| day_condition未実装 | "月,水,金"パース対応 |
| band判定が簡易(12:00基準のみ) | TimeWindow準拠 |

### 処理フロー

```
1. expand_rules_to_daily(): date range + day_condition + bandで日別展開
2. generate_shadowing_visits():
   for each (trainee, date):
     mentor_visits = get_mentor_visits(mentor_id, date)
     filter by band
     if mentor has no visits: skip (day off)
     for each mentor_visit:
       check conflict with trainee's existing visits
       if conflict: try_adjust_trainee_visit (flexのみ)
       if ok: create trainee copy
3. apply_to_result_rows(): 結果に追加
```

### visit_id生成

`{base_vid}_T_{trainee_id}` (衝突時は_2, _3追加)

### 2名体制+mentor overlap

mentorが2名体制訪問に参加 + traineeが同行する場合:
- traineeは観察者として記録（`[2名体制観察]` tag付き）
- need_staffのカウントには含めない

## パイプライン統合ポイント

```
Phase 3:  Level0 + atomic_assign ← CoupledVisitManager
Phase 4:  enforce_atomicity      ← CoupledVisitManager
Phase 5:  GapPack (anchor)       ← CoupledVisitManager
Phase 6:  sync_all_pairs         ← CoupledVisitManager
Phase 8:  Level1 (atomic re-ins) ← CoupledVisitManager
Phase 9:  enforce + sync         ← CoupledVisitManager
Phase 10: Level3 (exclude)       ← CoupledVisitManager
Phase 14: mentor expansion       ← MentorPairProcessor (後処理)
```
