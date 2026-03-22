# 01. 全体アーキテクチャ・パイプライン再設計

## 4層パイプライン

```
Layer 1: PREPARE (前処理)    → 10種データ統合・矛盾解決
Layer 2: MODEL (制約モデル)   → ConstraintChecker + ScoringEngine + TimeSlotManager
Layer 3: SOLVE (最適化)       → Phase A(固定枠) → Phase B(Greedy+バックトラック) → Phase C(局所探索)
Layer 4: EMIT (後処理)        → 検証・MentorPair展開・診断レポート
```

## Layer 1: PREPARE

### Step 1-1: DataLoad & Parse
- 10種のJSON入力を型付きPythonオブジェクトに変換
- 時刻→分変換、曜日セット化、緯度経度バリデーション

### Step 1-2: Normalize & Enrich
- weekly_patterns → 患者の時間枠上書き
- patient_changes適用: 時間変更→上書き、追加→挿入、キャンセル→除外
- special_week適用: ADD→追加、REPLACE→全置換
- confirmed_history → 患者ごと直近N週の担当スタッフ頻度マップ構築

### Step 1-3: Merge & Resolve
- 同一patient_id+dateの競合解決
- mentor_pairsからの暗黙的need_staff=2強制
- 指定スタッフIDが存在しない場合の除去

### Step 1-4: Validate & Report
- 必須フィールド、値域、参照整合性チェック
- 致命的エラー vs 警告の分類

## Layer 2: MODEL

### ConstraintModel
- HardConstraints H1-H11: 違反時は割当不可
- SoftConstraints S1-S7: ペナルティスコアで評価

### StaffTimeline (SortedList区間木)
- スタッフ×日ごとの時間管理
- O(log n) insert/remove/query
- GAS版のstaffDateMap（フラット配列）を置き換え

## Layer 3: SOLVE

### Phase A: FixedSlot Assign
- イベント→blocked_intervalsに登録
- time_type=固定 + specified_type=必須 → 候補1名で確定
- mentor_pairsの同行訪問 → ペアで同時確定

### Phase B: Greedy + Backtrack
- 優先度ソート（日付→needStaff降順→候補数昇順→ローテ→時刻）
- 全制約を統一チェック（ConstraintChecker.check_all）
- 2名体制はペア単位アトミック割当
- バックトラック（深さ3）で局所最適脱出
- StaffTimelineの空き区間に直接挿入 → GapPack不要

### Phase C: LocalSearch Optimize
- 2-opt（時間窓制約付き）+ Or-opt
- スタッフ間移動（訪問XをスタッフAからBに移動）
- 時間微調整

## Layer 4: EMIT

### Step 4-1: Verify All Constraints
- 全H1-H11を最終結果に再検証
- 違反0 = エンジン品質保証

### Step 4-2: Format Output
- GAS Bridge互換JSON変換
- 移動距離計算

### Step 4-3: Diagnose & Report
- 未割当原因分析
- スタッフ別負荷レポート
- 品質スコア(0-100)

## GAS版から廃止される処理

| 廃止処理 | 理由 |
|---------|------|
| GapPack | TimeSlotManagerで割当時に最適位置決定 |
| OverlapFix | TimeSlotManagerが重複挿入を拒否 |
| syncCoupledVisitTimes (5回呼出) | ペア単位アトミック割当で不要 |

## データフロー図

```
入力データ10種          Layer 1       Layer 2        Layer 3       Layer 4
                       PREPARE       MODEL          SOLVE         EMIT

staff_masters ─────── Parse ──── H1-H3,Timeline ── A,B,C ──── Format
patient_masters ───── Parse ──── S4,S7 ─────────── B,C ────── Format
weekly_requests ───── Merge ──── (EffectiveReq) ── B ──────── Format
events ────────────── Parse ──── blocked ────────── A ──────── Format
staff_changes ─────── Parse ──── H3,H4 ─────────── A,B ────── Diagnose
patient_changes ───── Merge ──────────────────────────────── Diagnose
special_week ──────── Merge ──────────────────────────────── Diagnose
mentor_pairs ──────── Resolve ── H10 ───────────── A ──────── MentorExpand
weekly_patterns ───── Enrich ──────────────────────────────── Diagnose
confirmed_history ─── Aggregate── S2(Rotation) ── B ──────── Diagnose
```

## 定数一覧

| 定数名 | 値 | 用途 |
|--------|-----|------|
| ASSIGN_BUFFER_MIN | 5分 | 訪問間の最小バッファ |
| EXTRA_BUFFER_MIN | 15分 | イベント前後のバッファ |
| EARTH_RADIUS_KM | 6371 | Haversine距離計算 |
| TWO_OPT_MAX_ITER | 200 | 2-opt最大反復 |
| BACKTRACK_DEPTH | 3 | バックトラック最大深さ |
| STAFF_SWAP_MAX_ITER | 100 | スタッフ間移動の最大反復 |
| ROTATION_LOOKBACK_WEEKS | 4 | ローテーション参照週数 |
| DEFAULT_AM | 540-720 | 午前デフォルト |
| DEFAULT_PM | 780-1020 | 午後デフォルト |
| DEFAULT_ALLDAY | 540-1080 | 終日デフォルト |

## モジュール構成

```
allocation_engine/
├── prepare/           # Layer 1
│   ├── parser.py
│   ├── request_merger.py
│   ├── staff_availability.py
│   └── patient_profile.py
├── model/             # Layer 2
│   ├── constraints.py
│   ├── scorer.py
│   ├── relaxation.py
│   ├── timeline.py
│   └── rotation_tracker.py
├── solve/             # Layer 3
│   ├── phase_a_fixed.py
│   ├── phase_b_greedy.py
│   ├── phase_c_optimize.py
│   ├── coupled_manager.py
│   └── mentor_processor.py
├── emit/              # Layer 4
│   ├── verifier.py
│   ├── formatter.py
│   └── diagnostics.py
└── constants.py
```
