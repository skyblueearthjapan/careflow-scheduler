"""
test_special_week.py - Tests for the special_week (特別訪問週間) logic
in AllocationEngine._apply_special_week().
"""
import sys
import os
import pytest

# Ensure python_engine package is importable
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from python_engine.allocation_models import (
    Patient, Staff, VisitRequest, SpecialWeekHeader, SpecialWeekDetail,
)
from python_engine.allocation_engine import AllocationEngine


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _make_engine(
    special_week_headers=None,
    special_week_details=None,
    patient_map=None,
):
    """Create a minimal AllocationEngine with special_week data."""
    staff = [
        Staff(sid="S1", name="田中", gender="女性", work_days=["Mon", "Tue", "Wed", "Thu", "Fri"]),
        Staff(sid="S2", name="鈴木", gender="男性", work_days=["Mon", "Tue", "Wed", "Thu", "Fri"]),
    ]
    if patient_map is None:
        patient_map = {
            "P001": Patient(pid="P001", name="患者A", area="東区"),
            "P002": Patient(pid="P002", name="患者B", area="西区"),
        }
    return AllocationEngine(
        staff_list=staff,
        patient_map=patient_map,
        events=[],
        staff_changes=[],
        special_week_headers=special_week_headers or [],
        special_week_details=special_week_details or [],
    )


def _make_requests():
    """Create sample weekly visit requests."""
    return [
        VisitRequest(request_id="R1", date_str="2026/03/30", pid="P001", pname="患者A",
                     service_min=60, need_staff=1, time_type="午前"),
        VisitRequest(request_id="R2", date_str="2026/04/01", pid="P001", pname="患者A",
                     service_min=60, need_staff=1, time_type="午後"),
        VisitRequest(request_id="R3", date_str="2026/03/31", pid="P002", pname="患者B",
                     service_min=30, need_staff=1, time_type="終日"),
    ]


# ===========================================================================
# Test: ヘッダ/明細なし → リクエスト変化なし
# ===========================================================================
class TestSpecialWeekNoData:
    def test_no_special_week_returns_unchanged(self):
        engine = _make_engine()
        reqs = _make_requests()
        result = engine._apply_special_week(reqs)
        assert len(result) == 3
        assert [r.request_id for r in result] == ["R1", "R2", "R3"]


# ===========================================================================
# Test: ADDモード — 既存リクエストを残しつつ追加
# ===========================================================================
class TestSpecialWeekAdd:
    def test_add_mode_appends_requests(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_001", pid="P001", pname="患者A",
                week_start="2026/03/30", mode="ADD", reason="退院直後の集中ケア",
            )
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_001", pid="P001",
                date_str="2026/03/31", time_type="午前",
                earliest_min=540, latest_min=720, service_min=60, need_staff=1,
            ),
            SpecialWeekDetail(
                special_week_id="SW_001", pid="P001",
                date_str="2026/04/02", time_type="午後",
                earliest_min=780, latest_min=1020, service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()
        result = engine._apply_special_week(reqs)

        # 既存3件 + 追加2件 = 5件
        assert len(result) == 5
        # 既存リクエストが残っている
        original_ids = {r.request_id for r in result[:3]}
        assert original_ids == {"R1", "R2", "R3"}
        # 追加分のIDがSW_で始まる
        added = result[3:]
        assert all(r.request_id.startswith("SW_") for r in added)
        assert added[0].date_str == "2026/03/31"
        assert added[1].date_str == "2026/04/02"

    def test_add_mode_preserves_other_patient(self):
        """ADDモードは他の患者のリクエストに影響しない"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_002", pid="P001", pname="患者A",
                mode="ADD", reason="テスト",
            )
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_002", pid="P001",
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()
        result = engine._apply_special_week(reqs)

        p002_reqs = [r for r in result if r.pid == "P002"]
        assert len(p002_reqs) == 1
        assert p002_reqs[0].request_id == "R3"

    def test_add_note_includes_reason(self):
        """追加されたリクエストのnoteに理由が含まれる"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_003", pid="P001", pname="患者A",
                mode="ADD", reason="退院直後",
            )
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_003", pid="P001",
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week(_make_requests())
        added = [r for r in result if r.request_id.startswith("SW_")]
        assert len(added) == 1
        assert "退院直後" in added[0].note

    def test_add_note_without_reason(self):
        """理由が空の場合は「特別訪問週間」のみ"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_004", pid="P001", pname="患者A",
                mode="ADD", reason="",
            )
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_004", pid="P001",
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week(_make_requests())
        added = [r for r in result if r.request_id.startswith("SW_")]
        assert added[0].note == "特別訪問週間"


# ===========================================================================
# Test: REPLACEモード — 対象患者の既存リクエスト全削除して置換
# ===========================================================================
class TestSpecialWeekReplace:
    def test_replace_removes_existing_and_adds_new(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_010", pid="P001", pname="患者A",
                mode="REPLACE", reason="スケジュール全面変更",
            )
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_010", pid="P001",
                date_str="2026/03/30", time_type="固定",
                start_min=600, end_min=660, service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()  # P001: R1,R2, P002: R3
        result = engine._apply_special_week(reqs)

        # P001の既存R1,R2は削除、SW追加1件 + P002のR3が残る = 2件
        assert len(result) == 2
        pids = [r.pid for r in result]
        assert pids.count("P001") == 1  # 置換後の1件
        assert pids.count("P002") == 1  # 他患者は影響なし

        p001_req = [r for r in result if r.pid == "P001"][0]
        assert p001_req.request_id.startswith("SW_")
        assert p001_req.start_min == 600

    def test_replace_does_not_affect_other_patient(self):
        """REPLACEは対象患者以外に影響しない"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_011", pid="P001", pname="患者A",
                mode="REPLACE", reason="テスト",
            )
        ]
        details = []  # 明細なし = P001の訪問を全削除
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()
        result = engine._apply_special_week(reqs)

        # P001は全削除（明細0件）、P002のR3だけ残る
        assert len(result) == 1
        assert result[0].pid == "P002"
        assert result[0].request_id == "R3"


# ===========================================================================
# Test: 患者マスタからのarea補完
# ===========================================================================
class TestSpecialWeekAreaEnrichment:
    def test_area_from_patient_map(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_020", pid="P001", pname="患者A", mode="ADD"),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_020", pid="P001",
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week(_make_requests())
        added = [r for r in result if r.request_id.startswith("SW_")]
        assert added[0].area == "東区"

    def test_area_empty_for_unknown_patient(self):
        """患者マスタに存在しないPIDの場合、areaは空"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_021", pid="P999", pname="不明患者", mode="ADD"),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_021", pid="P999",
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week([])
        assert result[0].area == ""


# ===========================================================================
# Test: 複数患者の同時特別訪問週間
# ===========================================================================
class TestSpecialWeekMultiplePatients:
    def test_two_patients_add(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_030", pid="P001", pname="患者A",
                mode="ADD", reason="集中ケア",
            ),
            SpecialWeekHeader(
                special_week_id="SW_031", pid="P002", pname="患者B",
                mode="ADD", reason="退院準備",
            ),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_030", pid="P001",
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
            SpecialWeekDetail(
                special_week_id="SW_031", pid="P002",
                date_str="2026/04/01", service_min=30, need_staff=2,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()
        result = engine._apply_special_week(reqs)

        # 既存3件 + 追加2件 = 5件
        assert len(result) == 5

    def test_mixed_add_and_replace(self):
        """P001=REPLACE, P002=ADD の同時適用"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_040", pid="P001", pname="患者A",
                mode="REPLACE", reason="全面変更",
            ),
            SpecialWeekHeader(
                special_week_id="SW_041", pid="P002", pname="患者B",
                mode="ADD", reason="追加ケア",
            ),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_040", pid="P001",
                date_str="2026/03/30", service_min=60, need_staff=1,
            ),
            SpecialWeekDetail(
                special_week_id="SW_041", pid="P002",
                date_str="2026/04/01", service_min=30, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()  # P001: R1,R2, P002: R3
        result = engine._apply_special_week(reqs)

        # P001: REPLACE → R1,R2削除 + SW1件 = 1件
        # P002: ADDで既存R3残 + SW1件 = 2件
        # 合計 3件
        assert len(result) == 3
        p001 = [r for r in result if r.pid == "P001"]
        p002 = [r for r in result if r.pid == "P002"]
        assert len(p001) == 1
        assert p001[0].request_id.startswith("SW_")
        assert len(p002) == 2


# ===========================================================================
# Test: 2名体制（need_staff=2）の特別訪問
# ===========================================================================
class TestSpecialWeekCoupled:
    def test_need_staff_2_preserved(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_050", pid="P001", pname="患者A",
                mode="ADD", reason="2名体制追加",
            ),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_050", pid="P001",
                date_str="2026/03/31", time_type="午前",
                earliest_min=540, latest_min=720,
                service_min=90, need_staff=2,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week(_make_requests())
        added = [r for r in result if r.request_id.startswith("SW_")]
        assert len(added) == 1
        assert added[0].need_staff == 2
        assert added[0].service_min == 90


# ===========================================================================
# Test: pidの補完（明細にpidがない場合はヘッダから取得）
# ===========================================================================
class TestSpecialWeekPidFallback:
    def test_detail_pid_empty_uses_header_pid(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_060", pid="P001", pname="患者A", mode="ADD"),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_060", pid="",  # 空
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week([])
        assert result[0].pid == "P001"

    def test_detail_pid_set_uses_detail_pid(self):
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_061", pid="P001", pname="患者A", mode="ADD"),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_061", pid="P002",  # 明細側で上書き
                date_str="2026/03/31", service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        result = engine._apply_special_week([])
        assert result[0].pid == "P002"


# ===========================================================================
# Test: allocate()統合テスト — special_weekが割当パイプラインを通るか
# ===========================================================================
class TestSpecialWeekIntegration:
    def test_allocate_with_special_week_add(self):
        """allocate()でADDモードの特別訪問が割り当てられる"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_100", pid="P001", pname="患者A",
                mode="ADD", reason="統合テスト",
            ),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_100", pid="P001",
                date_str="2026/03/30", time_type="午前",
                earliest_min=540, latest_min=720,
                service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()

        result = engine.allocate(reqs)

        assert "summary" in result
        assert "results" in result
        # 特別訪問のリクエストがnoteに含まれるか、またはP001の割当/未割当件数が増えているか
        p001_results = [r for r in result["results"] if r.pid == "P001" and not r.is_event]
        p001_unassigned = [u for u in result["unassigned"] if u.get("pid") == "P001"]
        sw_in_notes = any("特別訪問週間" in (r.note or "") for r in p001_results)
        # 通常2件(R1,R2) + 特別1件 = 3件以上あるはず
        total_p001 = len(p001_results) + len(p001_unassigned)
        assert sw_in_notes or total_p001 >= 3, \
            f"特別訪問が反映されていない: results={len(p001_results)}, unassigned={len(p001_unassigned)}"

    def test_allocate_with_special_week_replace(self):
        """allocate()でREPLACEモードが正しく通常リクエストを置換する"""
        headers = [
            SpecialWeekHeader(
                special_week_id="SW_101", pid="P001", pname="患者A",
                mode="REPLACE", reason="置換テスト",
            ),
        ]
        details = [
            SpecialWeekDetail(
                special_week_id="SW_101", pid="P001",
                date_str="2026/03/30", time_type="固定",
                start_min=600, end_min=660,
                service_min=60, need_staff=1,
            ),
        ]
        engine = _make_engine(special_week_headers=headers, special_week_details=details)
        reqs = _make_requests()

        result = engine.allocate(reqs)

        # P001の通常リクエスト(R1,R2)は割当結果に入らないはず
        all_visit_ids = [r.visit_id for r in result["results"] if r.pid == "P001"]
        assert not any("R1" in vid or "R2" in vid for vid in all_visit_ids), \
            "REPLACEなのにP001の通常リクエストが割当結果に含まれている"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
