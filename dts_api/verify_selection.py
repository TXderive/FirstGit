"""验证 Excel 筛选流程；只加载纯数据处理方法，不调用 FTP、Dify 或通知接口。"""

import ast
import logging
import tempfile
import unittest
from pathlib import Path

import pandas as pd


SOURCE = Path(__file__).with_name("auto_submit.py")
tree = ast.parse(SOURCE.read_text())
service = next(
    node for node in tree.body
    if isinstance(node, ast.ClassDef) and node.name == "DtsAutoSubmitService"
)
methods = {"_read_pending", "_cell", "_build_items", "_deduplicate", "_precheck"}
service.body = [node for node in service.body if getattr(node, "name", "") in methods]
module = ast.Module(body=[service], type_ignores=[])
logger = logging.getLogger("django")
logger.addHandler(logging.NullHandler())
logger.propagate = False
scope = {"pd": pd, "logger": logger}
exec(compile(ast.fix_missing_locations(module), str(SOURCE), "exec"), scope)
SelectionService = scope["DtsAutoSubmitService"]


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
        self.service.version_map = {"version-a": "P1"}

    def workbook(self, rows, reviews=None):
        columns = list(candidate())
        columns += sorted({key for row in rows for key in row} - set(columns))
        with pd.ExcelWriter(self.path, engine="openpyxl") as writer:
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
