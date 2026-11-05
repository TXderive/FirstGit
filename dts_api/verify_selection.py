"""使用临时 Excel 和本地服务替身验证筛选及预检查，不连接真实业务服务。"""

import ast
import logging
import tempfile
import unittest
from contextlib import nullcontext
from pathlib import Path, PurePosixPath
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pandas as pd


SOURCE = Path(__file__).with_name("auto_submit.py")
tree = ast.parse(SOURCE.read_text())
service = next(
    node for node in tree.body
    if isinstance(node, ast.ClassDef) and node.name == "DtsAutoSubmitService"
)
methods = {
    "_read_pending", "_cell", "_build_items", "_deduplicate", "_precheck",
    "_split_shared_baselines", "_log_step", "_split_history",
    "_get_proof_files", "precheck", "submit_all",
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
scope = {"pd": pd, "logger": logger, "Path": Path, "PurePosixPath": PurePosixPath}
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


class SelectionTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "version-a.xlsx"
        self.service = SelectionService()
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

    def run_precheck(self, versions=SHARED_VERSIONS):
        self.service.temp_dir = SimpleNamespace(name=self.directory.name)

        def download(ftp, remote, local_dir):
            local = Path(local_dir) / PurePosixPath(remote).name
            return local if local.is_file() else None

        self.service._download = Mock(side_effect=download)
        self.service._get_token = Mock(return_value="test-token")
        self.service.upload_proof = Mock()
        self.service.submit_workflow = Mock()
        with patch.dict(scope, {
            "FtpMgr": SimpleNamespace(
                get_ias_ftp_client=lambda: nullcontext(object()),
            ),
            "excel_list": [f"{version}.xlsx" for version in versions],
            "excel_base_url": "/baselines",
            "msg_base_url": "/minutes",
        }):
            return self.service.precheck()

    def assert_shared_submission_blocked(self, result, reason):
        self.assertEqual(result["ready"], [])
        self.assertEqual(len(result["blocked"]), 1)
        self.assertIn(reason, ";".join(result["blocked"][0]["errors"]))
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
