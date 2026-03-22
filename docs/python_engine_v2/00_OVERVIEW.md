# Python割当エンジン v2 設計書

## 概要

GAS移植版の8段階逐次パイプラインを、Pythonの強みを活かした4層アーキテクチャに再設計する。
10種全入力データの完全活用、監査で判明した5件の重大バグ根絶、GAS 6分制限からの解放による高品質アルゴリズムの導入を同時に達成する。

## 設計書一覧

| # | ファイル | 設計領域 | 担当 |
|---|---------|---------|------|
| 01 | `01_ARCHITECTURE.md` | 全体アーキテクチャ・パイプライン再設計 | チーフアーキテクト |
| 02 | `02_DATA_PREPROCESSING.md` | データ前処理・統合レイヤー | データ前処理設計 |
| 03 | `03_CONSTRAINT_ENGINE.md` | 制約エンジン・フィルタリング | 制約エンジン設計 |
| 04 | `04_TIME_MANAGEMENT.md` | 時間スロット管理・GapPack・ルート最適化 | 時間管理設計 |
| 05 | `05_ROTATION_TRACKER.md` | ローテーション・継続性ロジック | ローテーション設計 |
| 06 | `06_COUPLED_MENTOR.md` | 2名体制・同行研修の統合 | 2名体制・同行研修設計 |
| 07 | `07_OPTIMIZATION.md` | 最適化アルゴリズム高度化 | 最適化アルゴリズム設計 |
| 08 | `08_QUALITY_TESTING.md` | API・テスト・品質保証 | 品質保証・テスト設計 |

## 新アーキテクチャ: 4層パイプライン

```
Layer 1: PREPARE (前処理)    → 10種データ統合・矛盾解決
Layer 2: MODEL (制約モデル)   → ConstraintChecker + ScoringEngine + TimeSlotManager
Layer 3: SOLVE (最適化)       → Phase A(固定枠) → Phase B(Greedy+バックトラック) → Phase C(2-opt+SA)
Layer 4: EMIT (後処理)        → 検証・MentorPair展開・診断レポート
```

## GAS版から根本的に改善される5つのポイント

| # | GAS版の問題 | Python v2の解決策 |
|---|------------|-----------------|
| 1 | 制約チェックが3箇所に散在、Level1でNG/性別欠落 | ConstraintChecker で全制約を1クラスに統合 |
| 2 | GapPackでunregister忘れ、バッファ二重適用 | TimeSlotManager (SortedList基盤) が唯一の時間管理者 |
| 3 | 2名体制を個別割当→事後sync(5回呼出) | CoupledVisitManager でペア単位アトミック割当 |
| 4 | 5種データ未使用 | 前処理レイヤーでリクエスト統合 |
| 5 | 6分制限でGreedy+2-optのみ | プラグイン型ソルバー (Greedy改善→SA→CP-SAT) |

## モジュール構成

```
allocation_engine/
├── prepare/           # Layer 1: データ前処理
│   ├── parser.py
│   ├── request_merger.py
│   ├── staff_availability.py
│   └── patient_profile.py
├── model/             # Layer 2: 制約モデル
│   ├── constraints.py     # ConstraintChecker (H1-H11)
│   ├── scorer.py          # ScoringEngine (S1-S7)
│   ├── relaxation.py      # RelaxationPolicy
│   ├── timeline.py        # TimeSlotManager
│   └── rotation_tracker.py
├── solve/             # Layer 3: 最適化
│   ├── phase_a_fixed.py
│   ├── phase_b_greedy.py
│   ├── phase_c_optimize.py
│   ├── coupled_manager.py
│   └── mentor_processor.py
├── emit/              # Layer 4: 後処理
│   ├── verifier.py
│   ├── formatter.py
│   └── diagnostics.py
└── constants.py
```

## 段階的実装ロードマップ

| Phase | 内容 | 効果 |
|-------|------|------|
| Phase 1 | ConstraintChecker統一 + TimeSlotManager + 5重大バグ修正 | バグ根絶、制約違反ゼロ保証 |
| Phase 2 | 前処理レイヤー + RotationTracker | 全10データ活用 |
| Phase 3 | プラグイン型ソルバー (Greedy改善→SA→CP-SAT) | 未割当削減、距離改善 |

## 監査で判明した重大バグ（Phase 1で修正）

| # | 問題 | 深刻度 |
|---|------|--------|
| 1 | Level 1でNGスタッフ・性別制限が欠落 | 高 |
| 2 | GapPackでunregister_assignment忘れ | 高 |
| 3 | Level 3で時間窓制約(午前/午後等)が無視 | 高 |
| 4 | Level 0直後にcoupled atomicityチェックなし | 高 |
| 5 | maxPerDay: Level 0はsoft_cap, Level 1はmax_per_day（不整合） | 高 |
