# 02. データ前処理・統合レイヤー設計

## Raw入力型（10種のdataclass）

| # | 型名 | 元データ | 主要フィールド |
|---|------|---------|--------------|
| 1 | RawStaffMaster | staff_masters | staff_id, staff_name, gender, lat/lng, shift, work_days, max_per_day, alloc_pref |
| 2 | RawPatientMaster | patient_masters | patient_id, patient_name, area, lat/lng, service_minutes |
| 3 | RawWeeklyRequest | weekly_requests | request_id, date, weekday, patient_id, need_staff, specified_staff, ng_staff, sex_limit, cont_pref, time_type |
| 4 | RawEvent | events | event_id, staff_id, date, start/end, duration, fixed_slot |
| 5 | RawStaffChange | staff_changes | staff_id, date, restriction_type, start/end |
| 6 | RawWeeklyPattern | weekly_patterns | patient_id, day_code, start/end, service_min, need_staff |
| 7 | RawConfirmedHistory | confirmed_history | week_start, date, patient_id, staff_id |
| 8 | RawPatientChange | patient_changes | patient_id, date, operation, time_type, start/end, specified_staff, ng_staff |
| 9 | RawSpecialWeekHeader | special_week.headers | special_week_id, patient_id, week_start, mode(ADD/REPLACE) |
| 10 | RawSpecialWeekDetail | special_week.details | special_week_id, patient_id, date, time_type, start/end, service_min |
| 11 | RawMentorPair | mentor_pairs | trainee_staff_id, mentor_staff_id, start/end_date, band, day_condition, priority |

## Normalized出力型

- **EffectiveRequest**: 統合済みリクエスト（source: WEEKLY_REQUEST/PATIENT_CHANGE_ADD/SPECIAL_WEEK）
- **StaffDayAvailability**: スタッフ×日の可用性（is_off, available_intervals, blocked_intervals, max_visits）
- **PatientProfile**: 患者プロファイル（weekly_pattern, recent_staff_ids, dominant_staff, rotation_shift_base）
- **MentorPairExpanded**: 日別展開済み同行ペア
- **PreprocessedData**: 全出力をまとめたコンテナ

## リクエスト統合パイプライン

```
weekly_requests (97件)
    │
    ├─[1] special_week REPLACE → 該当patient_id+dateのリクエストを除外、明細で置換
    │
    ├─[2] special_week ADD → 明細リクエストを追加
    │
    ├─[3] patient_changes キャンセル → マッチするリクエストを除外
    │
    ├─[4] patient_changes 時間変更 → 時刻フィールドを上書き
    │
    ├─[5] patient_changes スタッフ変更 → 指定スタッフ/NGスタッフを上書き
    │
    └─[6] patient_changes 追加 → 新規EffectiveRequestを挿入
    │
    ▼
effective_requests
```

### 適用優先度

| 優先度 | データソース | 処理 |
|--------|-------------|------|
| 1 (最高) | special_week REPLACE | 通常リクエストを完全置換 |
| 2 | patient_changes キャンセル | 該当リクエスト除外 |
| 3 | patient_changes 時間/スタッフ変更 | フィールド上書き |
| 4 | patient_changes 追加 | 新規リクエスト挿入 |
| 5 | special_week ADD | 追加リクエスト挿入 |
| 6 | weekly_requests | ベースリクエスト |

change_policy="個別変更も適用"の場合、REPLACE後の明細にもpatient_changesを適用。

## スタッフ可用性マップ

```
staff_masters + staff_changes + events + mentor_pairs
    → staff_availability[staff_id][date] = StaffDayAvailability

処理:
1. staff_changes 休み → is_off=True
2. staff_changes 午前のみ → available=[shift_start, 720)
3. staff_changes 午後のみ → available=[780, shift_end)
4. staff_changes 時間制限 → availableから該当区間を差し引き(subtract_interval)
5. events → blocked_intervalsに追加、max_visits -= 1
6. mentor_pairs → is_mentor_day=True, mentor_for追記
```

## 患者プロファイル強化

```
patient_masters + weekly_patterns + confirmed_history
    → PatientProfile

処理:
1. weekly_patterns → day_code別にグルーピング
2. confirmed_history → 直近4週分のstaff_id出現頻度Counter
3. dominant_staff = Counter.most_common(1)
4. consecutive_weeks = 直近連続担当週数
```

## メンター同行展開

RawMentorPair → 日別展開:
- start_date〜end_dateの範囲内かつ今週に該当する日
- day_conditionが空でなければ曜日フィルタ（"月水金"→Mon,Wed,Fri）
- band→TimeInterval変換: 午前=[shift_start,720), 午後=[780,shift_end), 終日=全体

## エッジケース対処

| ケース | 対処 |
|--------|------|
| キャンセル対象が存在しない | 警告ログ、スキップ |
| REPLACE対象の患者にリクエストがない | 明細のみ挿入、警告 |
| 同一patient_id+dateに複数の時間変更 | 最後のchangeが勝つ |
| staff_changesで休み + eventsがある日 | 休みが優先 |
| mentorが休み | 同行行を生成しない |
| need_staff=2のリクエストへのキャンセル | 両枠とも除外 |
