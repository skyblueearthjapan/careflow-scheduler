# 08. API・テスト・品質保証設計

## APIレスポンス拡張

### 拡張フィールド

| フィールド | 内容 |
|-----------|------|
| quality_score | 品質スコア(0-100) + 5指標の内訳 |
| performance | 各フェーズの実行時間(ms) |
| constraint_report | ハード/ソフト制約違反レポート |
| explanations | 各訪問の選定理由（?detail=fullのみ） |

### エンドポイント

```
POST /api/allocate          # 割当実行 + 品質スコア
POST /api/allocate/debug    # デバッグ（件数+サンプル）
POST /api/allocate/compare  # GAS版との比較
GET  /api/allocate/health   # ヘルスチェック
```

### AllocationTracer（計装）

```python
class AllocationTracer:
    start_phase(phase_name)
    end_phase(phase_name)
    record_assignment(request_id, staff_id, selection_method, candidates, excluded)
    record_violation(type, severity, detail)
```

## 品質スコアリング (0-100)

### 5指標と重み

| 指標 | 重み | 計算方法 |
|------|------|---------|
| 未割当率 | 30% | 0%→100点, 5%→80点, 30%超→0点 |
| 距離効率 | 20% | 平均2km以下→100, 10km超→30 |
| 負荷バランス | 20% | CV(変動係数)=0→100, CV=1→0 |
| 希望充足率 | 20% | 指定必須/希望/継続性の充足度 |
| ローテーション品質 | 10% | ユニークスタッフ数/訪問回数 |

## テスト戦略

### テストファイル構成

```
tests/
  conftest.py                    # small_scenario(3x5), make_request
  fixtures/
    small_scenario.json
    real_data_97.json (匿名化)
    edge_cases/
  unit/
    test_constraints.py          # NG,性別,maxPerDay,勤務曜日,シフト,重複,変更
    test_scoring.py              # スコアリングロジック
    test_quality_score.py        # 品質スコア計算
  integration/
    test_small_scenario.py       # 基本割当,2名体制,負荷バランス,ローテーション,イベント
    test_edge_cases.py           # 0件,全員休み,2名体制のみ,同時刻固定
  regression/
    test_real_data.py            # 97件,品質>=70,ハード違反0,決定論性,5秒以内
  property/
    test_hard_constraints.py     # ハード制約は入力に関わらず違反しない
  api/
    test_api_endpoints.py
    test_compare_endpoint.py
```

### ユニットテスト例

| テスト | 検証内容 |
|--------|---------|
| test_ng_staff_excluded | NGリストのスタッフが候補に入らない |
| test_all_ng_empty_candidates | 全員NGなら候補は空 |
| test_female_only_excludes_male | 女性のみ→男性除外 |
| test_at_max_excluded | maxPerDay到達で除外 |
| test_non_work_day_excluded | 非勤務日で除外 |
| test_fixed_before_shift_excluded | シフト外の固定時刻で除外 |
| test_day_off_excluded | 休みの日は除外 |

### 回帰テスト（実データ97件）

```python
def test_quality_score_above_threshold(real_data):
    result = run_allocation_engine(real_data)
    assert result["quality_score"]["total"] >= 70

def test_no_hard_constraint_violations(real_data):
    assert len(result["constraint_report"]["hard_violations"]) == 0

def test_deterministic_results(real_data):
    result1 = run_allocation_engine(real_data)
    result2 = run_allocation_engine(real_data)
    assert result1 == result2

def test_performance_under_five_seconds(real_data):
    assert result["performance"]["total_ms"] < 5000
```

### プロパティテスト

全割当結果に対して:
1. NGスタッフに割当されていない
2. 性別制限を満たしている
3. 非勤務日に割当されていない
4. 2名体制で同一スタッフが重複していない

## /api/allocate/compare

GAS版とPython版の結果を並行比較。

### レスポンス

```json
{
  "comparison": {
    "summary": {
      "python_assigned": 85,
      "gas_assigned": 83,
      "matched": 72,
      "different_staff": 11,
      "python_only": 2,
      "match_rate_pct": 84.7
    },
    "quality_comparison": {
      "python_total": 82.5,
      "gas_total": 76.3,
      "python_better_dimensions": ["distance_efficiency"],
      "gas_better_dimensions": ["load_balance"]
    },
    "differences": [...]
  }
}
```

## 成功基準

- [ ] quality_score (0-100) がAPIレスポンスに含まれる
- [ ] performance.total_ms が各フェーズ別に報告される
- [ ] constraint_reportでハード制約違反が検出される
- [ ] ユニットテスト 20-30件通過
- [ ] 統合テスト 10ケース通過
- [ ] 回帰テスト: 実データ97件で品質スコア>=70
- [ ] プロパティテスト: ハード制約違反ゼロ（100回ランダム入力）
- [ ] /api/allocate/compare でGAS版との差異レポート取得可能
- [ ] 全テスト60秒以内
