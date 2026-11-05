"""使用临时 Excel 和本地服务替身验证筛选及预检查，不连接真实业务服务。"""

import ast
import logging
import shutil
import tempfile
import time
import unittest
from contextlib import nullcontext
from copy import deepcopy
from pathlib import Path, PurePosixPath
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd
from openpyxl import load_workbook


SOURCE = Path(__file__).with_name("auto_submit.py")
tree = ast.parse(SOURCE.read_text())
service = next(
    node for node in tree.body
    if isinstance(node, ast.ClassDef) and node.name == "DtsAutoSubmitService"
)
methods = {
    "_read_pending", "_cell", "_build_items", "_deduplicate", "_precheck",
    "_get_shared_baseline_groups", "_split_shared_baselines", "_log_step", "_split_history",
    "_get_proof_files", "precheck", "submit_all", "_download",
    "_history_key", "_get_history", "_save_history", "_set_excel_synced",
    "_writeback_targets", "_update_excel", "update_excel_submit_status", "run_auto_submit",
}
service.body = [node for node in service.body if getattr(node, "name", "") in methods]
constants = [
    node for node in tree.body
    if isinstance(node, ast.Assign)
    and any(isinstance(target, ast.Name) and target.id == "SHARED_BASELINE_GROUPS"
            for target in node.targets)
]
module = ast.Module(body=constants + [service], type_ignores=[])
logger = logging.getLogger("django")
logger.addHandler(logging.NullHandler())
logger.propagate = False
scope = {
    "pd": pd, "logger": logger, "Path": Path, "PurePosixPath": PurePosixPath,
    "time": time, "load_workbook": load_workbook,
}
exec(compile(ast.fix_missing_locations(module), str(SOURCE), "exec"), scope)
SelectionService = scope["DtsAutoSubmitService"]
SHARED_VERSIONS = ("aos5.0-master-1951", "aos5.0-master-1952")


def candidate(hwpsirt="H1", cve="CVE-2026-1001", **overrides):
    row = {
        "HWPSIRT ID": hwpsirt,
        "CVE ID": cve,
        "最新状态": "无影响",
        "已走单": "",
        "是否评审": "是",
    }
    row.update(overrides)
    return row


def review(hwpsirt="H1", cve="CVE-2026-1001", proof="proof.msg", **overrides):
    row = {"HWPSIRT ID": hwpsirt, "CVE ID": cve, "原始CVSS评分": proof}
    row.update(overrides)
    return row


class MemoryHistory:
    def __init__(self):
        self.records = {}

    def find_one(self, query):
        record = self.records.get((query["pbild"], query["hwpsirt_id"]))
        if record is not None and all(record.get(key) == value for key, value in query.items()):
            return deepcopy(record)
        return None

    def update_one(self, query, update, upsert=False):
        key = (query["pbild"], query["hwpsirt_id"])
        if key not in self.records:
            if not upsert:
                return
            self.records[key] = dict(query)
        self.records[key].update(deepcopy(update["$set"]))


class LocalFtp:
    def __init__(self, directory, failed_uploads=0):
        self.directory = Path(directory)
        self.failed_uploads = failed_uploads
        self.uploads = 0

    def download_file(self, local, remote):
        source = self.directory / PurePosixPath(remote).name
        if not source.is_file():
            return False
        shutil.copyfile(source, local)
        return True

    def upload_file(self, local, remote, force_upd=False):
        self.uploads += 1
        if self.failed_uploads:
            self.failed_uploads -= 1
            return False
        shutil.copyfile(local, self.directory / PurePosixPath(remote).name)
        return True


class SelectionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "version-a.xlsx"
        self.service = SelectionService()
        self.history = MemoryHistory()
        self.service._history = Mock(return_value=self.history)
        self.service.version_map = {
            "version-a": "P1", **{version: "P5" for version in SHARED_VERSIONS},
        }

    def workbook(self, rows, reviews=None, path=None):
        columns = list(candidate())
        columns += sorted({key for row in rows for key in row} - set(columns))
        with pd.ExcelWriter(path or self.path, engine="openpyxl") as writer:
            pd.DataFrame(rows, columns=columns).to_excel(
                writer, sheet_name="Sheet1", index=False,
            )
            if reviews is not None:
                review_columns = list(review())
                review_columns += sorted(
                    {key for row in reviews for key in row} - set(review_columns)
                )
                pd.DataFrame(reviews, columns=review_columns).to_excel(
                    writer, sheet_name="评审总表", index=False,
                )

    def read(self):
        return self.service._read_pending(self.path, "version-a")

    def test_only_unaffected_and_blank_submission_are_candidates(self):
        self.workbook(
            [
                candidate(),
                candidate("H2", **{"最新状态": "有影响"}),
                candidate("H3", **{"已走单": "是"}),
                candidate("H4", **{"已走单": "否"}),
            ],
            [review(hwpsirt=f"H{i}") for i in range(1, 5)],
        )
        result = self.read()
        self.assertEqual(result["HWPSIRT ID"].tolist(), ["H1"])
        self.assertEqual(result["qg_version"].tolist(), ["version-a"])

    def test_review_must_match_both_ids(self):
        self.workbook(
            [candidate("H1", "C1"), candidate("H2", "C2")],
            [review("H1", "C2"), review("H2", "C1")],
        )
        with self.assertLogs("django", level="WARNING") as logs:
            result = self.read()
        self.assertTrue(result.empty)
        self.assertEqual(len(logs.output), 2)
        self.assertTrue(all("评审核对失败" in line for line in logs.output))

    def test_review_only_records_cannot_create_candidates(self):
        self.workbook([candidate()], [review(), review("OTHER", "OTHER")])
        self.assertEqual(self.read()["HWPSIRT ID"].tolist(), ["H1"])

    def test_no_candidates_does_not_read_review_sheet(self):
        self.workbook([candidate(**{"最新状态": "有影响"})])
        result = self.read()
        self.assertTrue(result.empty)
        self.assertIn("proof_name", result.columns)

    def test_candidates_require_review_sheet(self):
        self.workbook([candidate()])
        with self.assertRaises(ValueError):
            self.read()

    def test_blank_ids_cannot_match(self):
        self.workbook(
            [candidate("", "C1"), candidate("H1", "")],
            [review("", "C1"), review("H1", "")],
        )
        self.assertTrue(self.read().empty)

    def test_existing_sheet1_reviewed_requirement_is_retained(self):
        self.workbook(
            [candidate(**{"是否评审": "否"})],
            [review(**{"是否评审": "是"})],
        )
        self.assertTrue(self.read().empty)

    def test_normalizes_whitespace(self):
        self.workbook(
            [candidate(" H1 ", " C1 ", **{"最新状态": " 无影响 ", "已走单": " ", "是否评审": " 是 "})],
            [review("H1", "C1", " proof.msg ")],
        )
        result = self.read()
        self.assertEqual(result["HWPSIRT ID"].tolist(), ["H1"])
        self.assertEqual(result["proof_name"].tolist(), ["proof.msg"])

    def test_proof_name_comes_from_review_sheet(self):
        self.workbook([candidate(proof_name="wrong.msg")], [review(proof="right.msg")])
        self.assertEqual(self.read()["proof_name"].tolist(), ["right.msg"])

    def test_missing_proof_is_blocked_by_existing_precheck(self):
        self.workbook([candidate()], [review(proof="")])
        items = self.service._build_items(self.read())
        ready, blocked = self.service._precheck(self.service._deduplicate(items))
        self.assertEqual(ready, [])
        self.assertEqual(len(blocked), 1)
        self.assertIn("缺少证明文件", blocked[0]["errors"])

    def test_conflicting_review_proofs_still_block_submission(self):
        self.workbook([candidate()], [review(proof="a.msg"), review(proof="b.msg")])
        items = self.service._build_items(self.read())
        ready, blocked = self.service._precheck(self.service._deduplicate(items))
        self.assertEqual(ready, [])
        self.assertIn("同一pbild+HWPSIRT对应多个证明文件", blocked[0]["errors"])

    def test_empty_review_sheet_produces_no_candidates(self):
        self.workbook([candidate()], [])
        self.assertTrue(self.read().empty)

    def shared_workbooks(self, rows_1951, rows_1952, review_1952=True):
        self.workbook(
            rows_1951, [review()],
            path=Path(self.directory.name) / f"{SHARED_VERSIONS[0]}.xlsx",
        )
        self.workbook(
            rows_1952, [review()] if review_1952 else None,
            path=Path(self.directory.name) / f"{SHARED_VERSIONS[1]}.xlsx",
        )
        (Path(self.directory.name) / "proof.msg").write_text("test meeting minutes")

    def run_precheck(self, versions=SHARED_VERSIONS, shared_groups=None):
        self.service.temp_dir = SimpleNamespace(name=self.directory.name)

        def download(ftp, remote, local_dir):
            local = Path(local_dir) / PurePosixPath(remote).name
            return local if local.is_file() else None

        self.service._download = Mock(side_effect=download)
        self.service._get_token = Mock(return_value="test-token")
        self.service.upload_proof = Mock()
        self.service.submit_workflow = Mock()
        settings = {
            "FtpMgr": SimpleNamespace(
                get_ias_ftp_client=lambda: nullcontext(object()),
            ),
            "excel_list": [f"{version}.xlsx" for version in versions],
            "excel_base_url": "/baselines",
            "msg_base_url": "/minutes",
        }
        if shared_groups is not None:
            settings["SHARED_BASELINE_GROUPS"] = shared_groups
        with patch.dict(scope, settings):
            return self.service.precheck()

    def assert_shared_submission_blocked(self, result, reason, count=1):
        self.assertEqual(result["ready"], [])
        self.assertEqual(len(result["blocked"]), count)
        for item in result["blocked"]:
            self.assertIn(reason, ";".join(item["errors"]))
        self.service._get_token.assert_not_called()
        self.assertFalse(any(
            PurePosixPath(call.args[1]).name == "proof.msg"
            for call in self.service._download.call_args_list
        ))
        self.assertEqual(self.service.submit_all(result), [])
        self.service.upload_proof.assert_not_called()
        self.service.submit_workflow.assert_not_called()

    def test_every_status_other_than_unaffected_blocks_shared_submission(self):
        for status in ("有影响", "待确认", "分析中", "", None, "无影响待确认"):
            with self.subTest(status=status):
                self.shared_workbooks(
                    [candidate()], [candidate(**{"最新状态": status})],
                    review_1952=False,
                )
                result = self.run_precheck()
                self.assert_shared_submission_blocked(result, "非无影响状态")
                self.assertIn(
                    SHARED_VERSIONS[1], ";".join(result["blocked"][0]["errors"]),
                )

    def test_shared_status_check_is_independent_of_file_order(self):
        self.shared_workbooks([candidate()], [candidate(**{"最新状态": "待确认"})])
        result = self.run_precheck(tuple(reversed(SHARED_VERSIONS)))
        self.assert_shared_submission_blocked(result, "非无影响状态")

    def test_non_unaffected_1951_also_blocks_1952(self):
        self.shared_workbooks([candidate(**{"最新状态": "待评审"})], [candidate()])
        result = self.run_precheck()
        self.assert_shared_submission_blocked(result, SHARED_VERSIONS[0])

    def test_both_unaffected_proceed_as_one_shared_submission(self):
        self.shared_workbooks([candidate()], [candidate()])
        result = self.run_precheck()
        self.assertEqual(result["blocked"], [])
        self.assertEqual(len(result["ready"]), 1)
        self.assertEqual(
            {item["qg_version"] for item in result["ready"][0]["targets"]},
            set(SHARED_VERSIONS),
        )
        self.assertTrue(Path(result["ready"][0]["proof_file"]).is_file())
        self.service._get_token.assert_called_once_with()

    def test_peer_status_is_checked_even_if_submitted_and_unreviewed(self):
        self.shared_workbooks(
            [candidate()],
            [candidate(**{"最新状态": "待确认", "已走单": "是", "是否评审": "否"})],
        )
        self.assert_shared_submission_blocked(self.run_precheck(), "非无影响状态")

    def test_missing_peer_workbook_blocks_shared_submission(self):
        self.shared_workbooks([candidate()], [candidate()])
        missing_file = f"{SHARED_VERSIONS[1]}.xlsx"
        (Path(self.directory.name) / missing_file).unlink()
        result = self.run_precheck()
        self.assert_shared_submission_blocked(result, "基线未成功读取")
        self.assertEqual(result["excel_failed"], [missing_file])

    def test_missing_peer_record_does_not_count_as_unaffected(self):
        self.shared_workbooks([candidate()], [candidate("OTHER", "OTHER")])
        self.assert_shared_submission_blocked(self.run_precheck(), "缺少对应HWPSIRT记录")

    def test_shared_status_uses_submission_hwpsirt_key_even_if_cve_differs(self):
        self.shared_workbooks(
            [candidate()], [candidate(cve="OTHER-CVE", **{"最新状态": "待确认"})],
        )
        self.assert_shared_submission_blocked(self.run_precheck(), "非无影响状态")

    def test_one_non_unaffected_duplicate_row_blocks_the_whole_group(self):
        self.shared_workbooks(
            [candidate()], [candidate(), candidate(**{"最新状态": "待确认"})],
        )
        self.assert_shared_submission_blocked(self.run_precheck(), "非无影响状态")

    def test_unrelated_hwpsirt_status_does_not_block_matching_unaffected_record(self):
        self.shared_workbooks(
            [candidate()],
            [candidate(), candidate("OTHER", "OTHER", **{"最新状态": "待确认"})],
        )
        result = self.run_precheck()
        self.assertEqual(len(result["ready"]), 1)
        self.assertEqual(result["blocked"], [])

    def test_other_versions_do_not_require_the_1951_1952_pair(self):
        self.workbook([candidate()], [review()])
        (Path(self.directory.name) / "proof.msg").write_text("test meeting minutes")
        result = self.run_precheck(("version-a",))
        self.assertEqual(len(result["ready"]), 1)
        self.assertEqual(result["blocked"], [])

    def test_shared_statuses_are_reloaded_on_each_run(self):
        self.shared_workbooks([candidate()], [candidate(**{"最新状态": "待确认"})])
        self.assert_shared_submission_blocked(self.run_precheck(), "非无影响状态")
        self.shared_workbooks([candidate()], [candidate()])
        result = self.run_precheck()
        self.assertEqual(len(result["ready"]), 1)
        self.assertEqual(result["blocked"], [])

    def baseline_workbooks(self, groups_by_pbild, overrides=None):
        overrides = overrides or {}
        versions = []
        for pbild, group in groups_by_pbild.items():
            for version in group:
                versions.append(version)
                self.service.version_map[version] = pbild
                self.workbook(
                    [candidate(**overrides.get(version, {}))], [review()],
                    path=Path(self.directory.name) / f"{version}.xlsx",
                )
        (Path(self.directory.name) / "proof.msg").write_text("test meeting minutes")
        return tuple(versions)

    def test_automatic_three_file_group_is_mutual_from_every_member(self):
        versions = ("baseline-a", "baseline-b", "baseline-c")
        for changed in versions:
            with self.subTest(non_unaffected=changed):
                self.baseline_workbooks(
                    {"P5": versions}, {changed: {"最新状态": "待确认"}},
                )
                result = self.run_precheck(versions, shared_groups=())
                self.assert_shared_submission_blocked(result, f"{changed}=待确认")

    def test_automatic_three_file_group_all_unaffected_can_proceed(self):
        versions = self.baseline_workbooks({"P5": ("baseline-a", "baseline-b", "baseline-c")})
        result = self.run_precheck(versions, shared_groups=())
        self.assertEqual(result["blocked"], [])
        self.assertEqual(len(result["ready"]), 1)
        self.assertEqual(
            {target["qg_version"] for target in result["ready"][0]["targets"]},
            set(versions),
        )

    def test_automatic_groups_are_independent_for_the_same_hwpsirt(self):
        versions = self.baseline_workbooks(
            {"P5": ("baseline-a", "baseline-b", "baseline-c"),
             "P6": ("baseline-d", "baseline-e")},
            {"baseline-c": {"最新状态": "待确认"}},
        )
        result = self.run_precheck(versions, shared_groups=())
        self.assertEqual(len(result["blocked"]), 1)
        self.assertEqual(result["blocked"][0]["pbild"], "P5")
        self.assertEqual(len(result["ready"]), 1)
        self.assertEqual(result["ready"][0]["pbild"], "P6")
        self.assertEqual(
            {target["qg_version"] for target in result["ready"][0]["targets"]},
            {"baseline-d", "baseline-e"},
        )

    def test_missing_download_does_not_remove_an_automatic_group_member(self):
        versions = self.baseline_workbooks({"P5": ("baseline-a", "baseline-b", "baseline-c")})
        (Path(self.directory.name) / "baseline-c.xlsx").unlink()
        result = self.run_precheck(versions, shared_groups=())
        self.assert_shared_submission_blocked(result, "基线未成功读取")
        self.assertEqual(result["excel_failed"], ["baseline-c.xlsx"])

    def test_unconfigured_version_map_entries_are_not_automatic_dependencies(self):
        versions = self.baseline_workbooks({"P5": ("baseline-a", "baseline-b")})
        self.service.version_map["not-in-excel-list"] = "P5"
        result = self.run_precheck(versions, shared_groups=())
        self.assertEqual(result["blocked"], [])
        self.assertEqual(len(result["ready"]), 1)

    def test_explicit_three_file_group_is_mutual_from_every_member(self):
        groups_by_pbild = {"PA": ("baseline-a",), "PB": ("baseline-b",), "PC": ("baseline-c",)}
        versions = ("baseline-a", "baseline-b", "baseline-c")
        for changed in versions:
            with self.subTest(non_unaffected=changed):
                self.baseline_workbooks(
                    groups_by_pbild, {changed: {"最新状态": "待确认"}},
                )
                result = self.run_precheck(versions, shared_groups=(versions,))
                self.assert_shared_submission_blocked(result, f"{changed}=待确认", count=2)

    def test_multiple_explicit_groups_are_independent(self):
        versions = self.baseline_workbooks(
            {"PA": ("baseline-a",), "PB": ("baseline-b",), "PC": ("baseline-c",),
             "PD": ("baseline-d",), "PE": ("baseline-e",)},
            {"baseline-c": {"最新状态": "待确认"}},
        )
        result = self.run_precheck(versions, shared_groups=(versions[:3], versions[3:]))
        self.assertEqual(len(result["blocked"]), 2)
        self.assertEqual(
            {item["qg_version"] for item in result["ready"]},
            {"baseline-d", "baseline-e"},
        )

    def test_overlapping_groups_are_transitive_in_both_directions(self):
        versions = ("baseline-a", "baseline-b", "baseline-c")
        for changed in ("baseline-a", "baseline-c"):
            with self.subTest(non_unaffected=changed):
                self.baseline_workbooks(
                    {"PA": (versions[0],), "PB": (versions[1],), "PC": (versions[2],)},
                    {changed: {"最新状态": "待确认"}, "baseline-b": {"已走单": "是"}},
                )
                result = self.run_precheck(
                    versions, shared_groups=((versions[0], versions[1]), (versions[1], versions[2])),
                )
                self.assert_shared_submission_blocked(result, f"{changed}=待确认")

    def test_explicit_and_automatic_groups_are_merged(self):
        versions = self.baseline_workbooks(
            {"P5": ("baseline-a", "baseline-b"), "P6": ("baseline-c",)},
            {"baseline-b": {"已走单": "是"}, "baseline-c": {"最新状态": "待确认"}},
        )
        result = self.run_precheck(versions, shared_groups=(("baseline-b", "baseline-c"),))
        self.assert_shared_submission_blocked(result, "baseline-c=待确认")

    def test_missing_pbild_does_not_link_unrelated_versions(self):
        versions = self.baseline_workbooks({"": ("baseline-a", "baseline-b")})
        with patch.dict(scope, {
            "excel_list": [f"{version}.xlsx" for version in versions],
            "SHARED_BASELINE_GROUPS": (),
        }):
            self.assertEqual(self.service._get_shared_baseline_groups(), ())

    def successful_history(self, **overrides):
        record = {
            "pbild": "P1", "hwpsirt_id": "H1", "status": "success",
            "dts_nos": ["DTS-EXISTING"], "excel_synced": False,
            "submit_time": "2000-01-01 00:00:00",
            "targets": [{"qg_version": "old-version", "hwpsirt_id": "H1", "cve_id": "CVE-2026-1001"}],
        }
        record.update(overrides)
        self.history.records[(record["pbild"], record["hwpsirt_id"])] = record
        return record

    def test_success_history_skips_dts_even_if_excel_marker_was_lost(self):
        for synced in (False, True):
            with self.subTest(excel_synced=synced):
                self.workbook([candidate()], [review()])
                self.successful_history(excel_synced=synced)
                result = self.run_precheck(("version-a",), shared_groups=())
                self.assertEqual(result["ready"], [])
                self.assertEqual(result["blocked"], [])
                self.assertEqual(len(result["history_hits"]), 1)
                hit = result["history_hits"][0]
                self.assertEqual(hit["status"], "already_submitted")
                self.assertEqual(hit["dts_nos"], ["DTS-EXISTING"])
                self.assertEqual({item["qg_version"] for item in hit["targets"]}, {"version-a"})
                self.service._get_token.assert_not_called()
                self.assertFalse(any(
                    PurePosixPath(call.args[1]).name == "proof.msg"
                    for call in self.service._download.call_args_list
                ))
                self.assertEqual(self.service.submit_all(result), [])
                self.service.upload_proof.assert_not_called()
                self.service.submit_workflow.assert_not_called()

    def test_failed_or_different_history_does_not_skip_new_submission(self):
        cases = (
            {"status": "submit_failed"}, {"pbild": "OTHER"}, {"hwpsirt_id": "OTHER"},
        )
        for changes in cases:
            with self.subTest(history=changes):
                self.history.records.clear()
                self.workbook([candidate()], [review()])
                (Path(self.directory.name) / "proof.msg").write_text("test minutes")
                self.successful_history(**changes)
                result = self.run_precheck(("version-a",), shared_groups=())
                self.assertEqual(len(result["ready"]), 1)
                self.assertEqual(result["history_hits"], [])

    def test_history_lookup_failure_cannot_fall_back_to_resubmission(self):
        self.workbook([candidate()], [review()])
        self.history.find_one = Mock(side_effect=RuntimeError("history unavailable"))
        with self.assertRaisesRegex(RuntimeError, "history unavailable"):
            self.run_precheck(("version-a",), shared_groups=())
        self.service._get_token.assert_not_called()
        self.service.submit_workflow.assert_not_called()

    def test_success_history_does_not_bypass_cross_baseline_status_check(self):
        self.shared_workbooks([candidate()], [candidate(**{"最新状态": "待确认"})])
        self.successful_history(pbild="P5")
        result = self.run_precheck()
        self.assert_shared_submission_blocked(result, "非无影响状态")
        self.assertEqual(result["history_hits"], [])

    def test_excel_writeback_marks_all_duplicate_rows_and_is_repeatable(self):
        self.workbook([candidate(), candidate(), candidate(**{"已走单": "是"})], [review()])
        targets = [{"hwpsirt_id": "H1", "cve_id": "CVE-2026-1001"}]
        self.assertEqual(self.service._update_excel(self.path, targets), (2, []))
        self.assertEqual(pd.read_excel(self.path, sheet_name="Sheet1")["已走单"].tolist(), ["是"] * 3)
        self.assertEqual(self.service._update_excel(self.path, targets), (0, []))

    def test_excel_missing_target_does_not_partially_mark_other_rows(self):
        self.workbook([candidate()], [review()])
        targets = [
            {"hwpsirt_id": "H1", "cve_id": "CVE-2026-1001"},
            {"hwpsirt_id": "MISSING", "cve_id": "CVE-2026-1001"},
        ]
        self.assertEqual(self.service._update_excel(self.path, targets), (0, [targets[1]]))
        marker = pd.read_excel(self.path, sheet_name="Sheet1")["已走单"].fillna("").tolist()
        self.assertEqual(marker, [""])

    def ftp_settings(self, ftp):
        return {
            "FtpMgr": SimpleNamespace(get_ias_ftp_client=lambda: nullcontext(ftp)),
            "excel_list": ["version-a.xlsx"], "excel_base_url": "/baselines",
            "msg_base_url": "/minutes", "SHARED_BASELINE_GROUPS": (),
        }

    def simulated_task(self, ftp, day):
        task = SelectionService()
        task.version_map = dict(self.service.version_map)
        local = Path(self.directory.name) / day
        local.mkdir()
        task.temp_dir = SimpleNamespace(name=str(local))
        task._history = Mock(return_value=self.history)
        task._get_token = Mock(return_value="test-token")
        task.build_submit_data = Mock(return_value={})
        task.upload_proof = Mock(return_value="test-file-id")
        task.submit_workflow = Mock(return_value={
            "status": "success", "dts_nos": ["DTS-NEW"], "error": "",
        })
        task._save_history = Mock(wraps=task._save_history)
        task._build_result_report = Mock(return_value="test report")
        task.send_summary = Mock()
        with patch.dict(scope, self.ftp_settings(ftp)):
            summary = task.run_auto_submit()
        return task, summary

    def test_successful_dts_with_repeated_ftp_failures_is_not_resubmitted_next_days(self):
        self.workbook([candidate(), candidate()], [review()])
        (Path(self.directory.name) / "proof.msg").write_text("test minutes")
        ftp = LocalFtp(self.directory.name, failed_uploads=2)
        first, first_summary = self.simulated_task(ftp, "day1")
        first.submit_workflow.assert_called_once()
        self.assertEqual(first_summary["success"], 1)
        self.assertEqual(first_summary["excel_failed"], 1)
        record = self.history.records[("P1", "H1")]
        self.assertEqual(record["status"], "success")
        self.assertFalse(record["excel_synced"])
        record["submit_time"] = "2000-01-01 00:00:00"

        second, second_summary = self.simulated_task(ftp, "day2")
        second.submit_workflow.assert_not_called()
        second.upload_proof.assert_not_called()
        second._get_token.assert_not_called()
        second._save_history.assert_not_called()
        self.assertEqual(second_summary["total"], 0)
        self.assertEqual(second_summary["history_hits"], 1)
        self.assertEqual(second_summary["excel_failed"], 1)
        self.assertFalse(record["excel_synced"])

        third, third_summary = self.simulated_task(ftp, "day3")
        third.submit_workflow.assert_not_called()
        self.assertEqual(third_summary["history_hits"], 1)
        self.assertEqual(third_summary["excel_updated"], 2)
        self.assertTrue(record["excel_synced"])
        self.assertEqual(record["submit_time"], "2000-01-01 00:00:00")
        self.assertEqual(pd.read_excel(self.path, sheet_name="Sheet1")["已走单"].tolist(), ["是", "是"])

    def test_writeback_retry_redownloads_after_failed_upload_in_same_process(self):
        self.workbook([candidate()], [review()])
        (Path(self.directory.name) / "proof.msg").write_text("test minutes")
        ftp = LocalFtp(self.directory.name, failed_uploads=1)
        task, _ = self.simulated_task(ftp, "day1")
        repair = {
            "pbild": "P1", "hwpsirt_id": "H1", "status": "already_submitted",
            "targets": [{"qg_version": "version-a", "hwpsirt_id": "H1", "cve_id": "CVE-2026-1001"}],
        }
        with patch.dict(scope, self.ftp_settings(ftp)):
            result = task.update_excel_submit_status([repair])
        self.assertEqual(result, {"updated": 1, "failed": 0})
        self.assertEqual(ftp.uploads, 2)
        self.assertEqual(pd.read_excel(self.path, sheet_name="Sheet1")["已走单"].tolist(), ["是"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
