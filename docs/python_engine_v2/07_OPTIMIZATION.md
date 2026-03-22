# 07. 最適化アルゴリズム高度化設計

## アルゴリズム比較表

| アルゴリズム | 解品質 | 実行時間 | 実装難易度 | 外部依存 | 推奨Phase |
|-------------|--------|---------|-----------|---------|----------|
| Greedy改善版 | C+ | <1秒 | 低 | なし | Phase 1 |
| SA(焼きなまし) | A | 5-15秒 | 中 | なし | Phase 2 |
| CP-SAT(OR-Tools) | A+ | 1-5秒 | 中 | OR-Tools | Phase 3 |
| GA(遺伝的) | B+ | 10-30秒 | 高 | なし | 不採用 |
| CSP | B | 不定 | 中 | python-constraint | 不採用 |

## 多目的最適化

```
Total_Cost = w1*F_unassigned + w2*F_distance + w3*F_balance + w4*F_preference

w1 = 10000 (未割当1件 = 最重要)
w2 = 1     (距離km)
w3 = 100   (負荷バランス分散)
w4 = 1     (希望ペナルティ合計)
```

## Phase 1: Greedy改善版

### 改善項目

| # | 改善 | 効果 |
|---|------|------|
| G1 | リクエスト処理順序最適化 | 未割当削減 |
| G2 | softCap段階的緩和（2パス） | 未割当削減 |
| G3 | Level1候補: Top10→全スタッフ | 未割当削減 |
| G4 | 入替え挿入（Ejection Chain） | 未割当削減 |
| G5 | 2-opt上限撤廃（50→収束まで） | 距離改善 |
| G6 | Or-opt追加（1訪問/2連続移動） | 距離改善 |
| G7 | 統合スコア関数 | バランス改善 |

### 入替え挿入（G4）

割当済みreq_Aを外し、req_Aのスタッフにreqをはめ、req_Aを別スタッフに再配置。

### Or-opt（G6）

1訪問または連続2訪問を抜き出して別位置に挿入し、距離改善なら採用。

## Phase 2: シミュレーテッドアニーリング

### 近傍操作

| 操作 | 確率 | 内容 |
|------|------|------|
| swap_staff | 35% | 2訪問のスタッフ交換 |
| move_to_staff | 30% | 1訪問を別スタッフに移動 |
| swap_time | 15% | 同スタッフ内で時刻交換 |
| unassign_rescue | 20% | 未割当を割当済みと入替え |

### パラメータ

```
T_START = 1000, T_END = 0.1, ALPHA = 0.995
MAX_ITER = 50000, TIME_LIMIT = 25秒
再加熱: T_START * 0.3
```

## Phase 3: CP-SAT (OR-Tools)

### 変数

```
x[r,s] = リクエストrをスタッフsに割り当てるか (bool)
t[r]   = リクエストrの開始時刻（分単位の整数変数）
u[r]   = リクエストrが未割当か (bool)
```

変数数: ~97 x 5 x 20 = ~9,700変数（数秒で解ける）

### 主要制約

- NoOverlap（CP-SATの強力な組込制約）
- ペア原子性: u[r1] == u[r2]
- maxPerDay/性別/NG/勤務曜日: x[r,s] = 0 で除外

## プラグイン型ソルバー

```python
class BaseSolver(ABC):
    def solve(self, problem) -> AllocationResult: ...
    def name(self) -> str: ...

class GreedySolver(BaseSolver): ...  # Phase 1
class SASolver(BaseSolver): ...      # Phase 2
class CPSATSolver(BaseSolver): ...   # Phase 3

# API: solver_nameパラメータで選択
POST /api/allocate { ..., "solver": "greedy" | "sa" | "cpsat" }
```

## 実行時間見積もり

| ソルバー | 97件/5名 | 200件/10名 | 500件/20名 |
|---------|----------|-----------|-----------|
| Greedy | <1秒 | 1-2秒 | 3-5秒 |
| SA | 5-10秒 | 10-20秒 | 20-30秒 |
| CP-SAT | 1-5秒 | 5-15秒 | 30-120秒 |
