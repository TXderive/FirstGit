# -*- coding: utf-8 -*-
import json
import logging
import time
from collections import defaultdict
from pathlib import Path, PurePosixPath
from tempfile import TemporaryDirectory

import pandas as pd
from openpyxl import load_workbook
from prettytable import PrettyTable

from test_system_framework.cve.utils.cve_mongo_handler import (
    CveMongoDbHandler,
)
from test_system_manager.setting.models.user import WeLinkToken
from test_system_server.utils.ftp_mgr import FtpMgr
from test_system_third_party.const import WELINK_NOTICE_NO4
from test_system_third_party.service.dts_api.backend.config_loader import cfg
from test_system_third_party.service.dts_api.backend.dify_helper import (
    run_workflow,
    upload_file_to_dify,
)
from test_system_third_party.service.dts_api.dts_const import (
    ASSESSORS,
    OPERATOR,
    UNAFFECTED_REASON,
    excel_base_url,
    excel_list,
    msg_base_url, )
from test_system_third_party.welink import WeLinkMsg

logger = logging.getLogger("django")
HISTORY_COL = "dts_auto_submit_history"
# 自动按 excel_list 中各版本的 pbild 分组；这里可补充显式关联组。
# 每组可含任意多个版本（不带 .xlsx），组内双向检查，重叠组会合并。
SHARED_BASELINE_GROUPS = (
    ("aos5.0-master-1951", "aos5.0-master-1952"),
)


class DtsAutoSubmitService:
    def __init__(self):
        self.temp_dir = TemporaryDirectory()
        self.version_map = self._get_version_pbild_map()
        self.excel_paths = {}

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.temp_dir.cleanup()

    @staticmethod
    def _log_step(step, title):
        logger.info("=" * 70)
        logger.info("Step %s. %s", step, title)
        logger.info("=" * 70)

    @staticmethod
    def _add_version_count(stats, item, key):
        targets = item.get("targets", [item])
        versions = {x.get("qg_version", "") for x in targets}
        for version in versions:
            if version:
                stats[version][key] += 1

    def _build_result_report(self, precheck, results, excel_result):
        stats = defaultdict(
            lambda: {
                "success": 0,
                "failed": 0,
                "blocked": 0,
                "history": 0,
            }
        )

        for item in results:
            key = "success" if item["status"] == "success" else "failed"
            self._add_version_count(stats, item, key)

        for item in precheck["blocked"]:
            self._add_version_count(stats, item, "blocked")

        for item in precheck["history_hits"]:
            self._add_version_count(stats, item, "history")

        table = PrettyTable()
        table.field_names = [
            "AOS版本",
            "走单成功",
            "走单失败",
            "预检查失败",
            "历史已走单",
        ]
        table.align = "l"

        for version in sorted(stats):
            data = stats[version]
            table.add_row([
                version,
                data["success"],
                data["failed"],
                data["blocked"],
                data["history"],
            ])

        table.add_row([
            "TOTAL",
            sum(x["success"] for x in stats.values()),
            sum(x["failed"] for x in stats.values()),
            sum(x["blocked"] for x in stats.values()),
            sum(x["history"] for x in stats.values()),
        ])

        success = sum(
            item["status"] == "success"
            for item in results
        )
        failed = len(results) - success
        summary_text = (
            f"本次共提交 {len(results)} 单，成功 {success} 单，"
            f"失败 {failed} 单，预检查失败 {len(precheck['blocked'])} 单。"
        )
        return (
            "【dts定时走单】成功！\n"
            f"{summary_text}\n\n"
            f"{table}\n\n"
        )

    @staticmethod
    def _cell(row, col):
        value = row.get(col)
        return "" if pd.isna(value) else str(value).strip()

    @staticmethod
    def _download(ftp, remote, local_dir):
        local = Path(local_dir) / PurePosixPath(remote).name
        return local if ftp.download_file(str(local), remote) else None

    @staticmethod
    def _get_version_pbild_map():
        col = CveMongoDbHandler.get_collection("version_info")
        data = col.find({}, {"_id": 0, "qgVersion": 1, "pbiId": 1})
        return {
            str(x["qgVersion"]).strip(): str(x["pbiId"]).strip()
            for x in data if x.get("qgVersion") and x.get("pbiId")
        }

    @staticmethod
    def _history():
        return CveMongoDbHandler.get_collection(HISTORY_COL)

    @staticmethod
    def _history_key(item):
        return {
            "pbild": item["pbild"],
            "hwpsirt_id": item["hwpsirt_id"],
        }

    def _get_history(self, item):
        if not item["pbild"] or not item["hwpsirt_id"]:
            return None
        return self._history().find_one({
            **self._history_key(item),
            "status": "success",
        })

    def _save_history(self, item, result):
        targets = [
            {
                "qg_version": x["qg_version"],
                "hwpsirt_id": x["hwpsirt_id"],
                "cve_id": x["cve_id"],
            }
            for x in item["targets"]
        ]
        data = {
            **self._history_key(item),
            "status": "success",
            "cve_ids": sorted({x["cve_id"] for x in item["targets"]}),
            "targets": targets,
            "dts_nos": result["dts_nos"],
            "proof_name": item["proof_name"],
            "excel_synced": False,
            "submit_time": time.strftime("%Y-%m-%d %H:%M:%S"),
        }
        self._history().update_one(
            self._history_key(item),
            {"$set": data},
            upsert=True,
        )

    def _set_excel_synced(self, item, synced):
        self._history().update_one(
            self._history_key(item),
            {"$set": {
                "excel_synced": synced,
                "excel_sync_time": time.strftime("%Y-%m-%d %H:%M:%S"),
            }},
        )

    def _read_pending(self, excel_file, qg_version, baseline_statuses=None):
        """先筛选 Sheet1 待走单项，再核对评审总表并取得证明文件名。"""
        keys = ["HWPSIRT ID", "CVE ID"]

        with pd.ExcelFile(excel_file) as workbook:
            data = pd.read_excel(workbook, sheet_name="Sheet1")
            for col in keys:
                data[col] = data[col].fillna("").astype(str).str.strip()

            status = data["最新状态"].fillna("").astype(str).str.strip()
            # 在筛选前保留完整状态：其他状态、已走单或未评审的记录也参与核对。
            if baseline_statuses is not None:
                baseline_statuses[qg_version] = {}
                for hwpsirt, value in zip(data["HWPSIRT ID"], status):
                    if hwpsirt:
                        baseline_statuses[qg_version].setdefault(
                            hwpsirt, set(),
                        ).add(value)

            submitted = data["已走单"].fillna("").astype(str).str.strip()
            pending = data.loc[
                (status == "无影响") & (submitted == "")
            ].copy()
            pending["qg_version"] = qg_version
            logger.info(
                "[Sheet1初筛] qg=%s total=%s pending=%s",
                qg_version, len(data), len(pending),
            )

            if pending.empty:
                pending["proof_name"] = ""
                return pending

            review = pd.read_excel(workbook, sheet_name="评审总表")[
                keys + ["原始CVSS评分"]
            ].copy()

        for col in keys:
            review[col] = review[col].fillna("").astype(str).str.strip()

        # 空 ID 不能作为二次核对的匹配依据。
        review = review.loc[review[keys].ne("").all(axis=1)].copy()

        review["原始CVSS评分"] = (
            review["原始CVSS评分"].fillna("").astype(str).str.strip()
        )
        review.rename(
            columns={"原始CVSS评分": "proof_name"},
            inplace=True,
        )

        # 以 Sheet1 候选为主表；证明文件名只取自评审总表。
        result = pending.drop(columns=["proof_name"], errors="ignore").merge(
            review,
            on=keys,
            how="left",
            indicator="_review_match",
        )
        matched = result["_review_match"] == "both"
        for _, row in result.loc[~matched].iterrows():
            logger.warning(
                "[评审核对失败] qg=%s hwpsirt=%s cve=%s "
                "reason=评审总表无有效的双ID匹配记录",
                qg_version, row["HWPSIRT ID"], row["CVE ID"],
            )

        # 保留原有的 Sheet1“是否评审=是”限制。
        reviewed = result["是否评审"].fillna("").astype(str).str.strip()
        result = result.loc[matched & (reviewed == "是")].copy()
        result.drop(columns=["_review_match"], inplace=True)
        logger.info(
            "[评审总表核对] qg=%s candidates=%s matched_and_reviewed=%s",
            qg_version, len(pending), len(result),
        )
        return result

    def _build_items(self, df):
        result = []
        for _, row in df.iterrows():
            qg = self._cell(row, "qg_version")
            result.append({
                "hwpsirt_id": self._cell(row, "HWPSIRT ID"),
                "cve_id": self._cell(row, "CVE ID"),
                "qg_version": qg,
                "pbild": self.version_map.get(qg, ""),
                "proof_name": self._cell(row, "proof_name"),
                "proof_file": "",
            })
        return result

    @staticmethod
    def _deduplicate(items):
        groups = {}
        for item in items:
            key = (item["pbild"], item["hwpsirt_id"])
            if key not in groups:
                groups[key] = {**item, "targets": [], "group_error": ""}
            groups[key]["targets"].append(item)

        for item in groups.values():
            cves = {x["cve_id"] for x in item["targets"]}
            proofs = {x["proof_name"] for x in item["targets"]}
            errors = []
            if len(cves) > 1:
                errors.append("同一pbild+HWPSIRT对应多个CVE")
            if len(proofs) > 1:
                errors.append("同一pbild+HWPSIRT对应多个证明文件")
            item["group_error"] = ";".join(errors)

            if len(item["targets"]) > 1:
                versions = sorted({x["qg_version"] for x in item["targets"]})
                logger.info("[走单去重] hwpsirt=%s pbild=%s versions=%s", item["hwpsirt_id"], item["pbild"], ",".join(versions))
        return list(groups.values())

    def _get_shared_baseline_groups(self):
        """合并相同 pbild 和显式配置的关联关系，得到互相核对的版本组。"""
        by_pbild = {}
        for file_name in excel_list:
            version = Path(file_name).stem
            pbild = self.version_map.get(version, "")
            if pbild:
                by_pbild.setdefault(pbild, set()).add(version)

        groups = [set(group) for group in SHARED_BASELINE_GROUPS]
        groups.extend(versions for versions in by_pbild.values() if len(versions) > 1)
        merged = []
        for group in groups:
            component, separate = set(group), []
            for existing in merged:
                if component.intersection(existing):
                    component.update(existing)
                else:
                    separate.append(existing)
            merged = separate + [component]
        return tuple(tuple(sorted(group)) for group in merged if len(group) > 1)

    def _split_shared_baselines(self, items, baseline_statuses):
        """共用 DTS 单据在每个基线都明确为无影响，才允许继续走单。"""
        groups = self._get_shared_baseline_groups()
        pending, blocked = [], []
        for item in items:
            versions = {x["qg_version"] for x in item["targets"]}
            conflicts, missing, unread = {}, set(), set()
            for group in groups:
                if not versions.intersection(group):
                    continue
                for version in group:
                    if version not in baseline_statuses:
                        unread.add(version)
                        continue
                    # DTS 接口按 pbild + HWPSIRT ID 提交，CVE 不拆成独立单据。
                    statuses = baseline_statuses[version].get(
                        item["hwpsirt_id"], set(),
                    )
                    if not statuses:
                        missing.add(version)
                    elif statuses != {"无影响"}:
                        conflicts[version] = statuses

            reasons = []
            if conflicts:
                reasons.append(
                    "共用DTS单据存在非无影响状态: " + ";".join(
                        f"{version}={','.join(value or '空白' for value in sorted(statuses))}"
                        for version, statuses in sorted(conflicts.items())
                    )
                )
            if missing:
                reasons.append(
                    "共用DTS基线缺少对应HWPSIRT记录，无法确认无影响: "
                    + ",".join(sorted(missing))
                )
            if unread:
                reasons.append(
                    "共用DTS基线未成功读取，无法核对影响状态: "
                    + ",".join(sorted(unread))
                )
            if reasons:
                item["errors"] = (
                    [item["group_error"]] if item["group_error"] else []
                ) + reasons
                blocked.append(item)
            else:
                pending.append(item)
        return pending, blocked

    def _split_history(self, items):
        """成功历史阻止再次提交；使用本轮目标补写尚未同步的 Excel。"""
        pending, history_hits = [], []
        for item in items:
            history = self._get_history(item)
            if history is None:
                pending.append(item)
                continue

            # 即使历史标记已同步，本轮仍筛出未走单行时也要补写；不复用旧 targets。
            result = {
                **item,
                "status": "already_submitted",
                "dts_nos": history.get("dts_nos", []),
                "error": "",
            }
            history_hits.append(result)
            logger.info(
                "[历史已走单] hwpsirt=%s pbild=%s dts=%s 仅补写Excel，不再提交",
                item["hwpsirt_id"], item["pbild"],
                ",".join(result["dts_nos"]),
            )
        return pending, history_hits

    def _get_proof_files(self, ftp, items):
        cache, missing = {}, []
        for item in items:
            name = item["proof_name"]
            if not name:
                missing.append(item)
                continue

            if name not in cache:
                remote = str(PurePosixPath(msg_base_url) / name)
                cache[name] = self._download(
                    ftp,
                    remote,
                    self.temp_dir.name,
                )

            if cache[name]:
                item["proof_file"] = str(cache[name])
            else:
                missing.append(item)
        return missing

    @staticmethod
    def _precheck(items):
        ready, blocked = [], []
        for item in items:
            errors = []
            if not item["hwpsirt_id"]:
                errors.append("HWPSIRT ID为空")
            if not item["pbild"]:
                errors.append("缺少pbild")
            if item["group_error"]:
                errors.append(item["group_error"])
            if not item["proof_file"]:
                errors.append("缺少证明文件")

            item["errors"] = errors
            (blocked if errors else ready).append(item)
        return ready, blocked

    @staticmethod
    def _get_token():
        resp = run_workflow(
            {"tokennum": 1},
            cfg("dify", "get_token_app_key"),
            user="dts_auto_submit",
        )
        token = resp.get("data", {}).get("outputs", {}).get("token", "")
        if not token:
            raise RuntimeError("获取W3 Token失败")
        return token

    def precheck(self):
        items, excel_failed = [], []
        baseline_statuses = {}
        self.excel_paths = {}
        self._log_step(2, "获取并筛选基线Excel")

        with FtpMgr.get_ias_ftp_client() as ftp:
            for file_name in excel_list:
                remote = str(PurePosixPath(excel_base_url) / file_name)
                excel = self._download(ftp, remote, self.temp_dir.name)
                if not excel:
                    excel_failed.append(file_name)
                    continue

                qg = Path(file_name).stem
                self.excel_paths[qg] = excel
                items.extend(self._build_items(
                    self._read_pending(excel, qg, baseline_statuses)
                ))

            logger.info("Excel总数: %s", len(excel_list))
            logger.info("Excel下载失败: %s", len(excel_failed))
            logger.info("筛选待走单: %s", len(items))

            grouped = self._deduplicate(items)
            grouped, baseline_blocked = self._split_shared_baselines(
                grouped, baseline_statuses,
            )
            logger.info("共用基线核对失败: %s", len(baseline_blocked))
            pending, history_hits = self._split_history(grouped)

            self._log_step(4, "检查证明文件")
            missing = self._get_proof_files(ftp, pending)
            missing_names = sorted({
                x["proof_name"] for x in missing if x["proof_name"]
            })
            logger.info("待真实检查: %s", len(pending))
            logger.info("历史成功: %s", len(history_hits))
            logger.info("缺失证明文件: %s", len(missing_names))

        self._log_step(5, "走单前预检查")
        ready, blocked = self._precheck(pending)
        blocked = baseline_blocked + blocked

        logger.info("可走单: %s", len(ready))
        logger.info("预检查失败: %s", len(blocked))
        for item in blocked:
            logger.warning(
                "[预检查失败] cve=%s hwpsirt=%s qg=%s reason=%s",
                item["cve_id"],
                item["hwpsirt_id"],
                item["qg_version"],
                ";".join(item["errors"]),
            )

        self._log_step(6, "获取一次Token")
        token = self._get_token() if ready else ""
        logger.info("Token: %s", "获取成功" if token else "跳过")

        return {
            "ready": ready,
            "blocked": blocked,
            "history_hits": history_hits,
            "missing_files": missing_names,
            "excel_failed": excel_failed,
            "token": token,
        }

    @staticmethod
    def upload_proof(proof_file):
        return upload_file_to_dify(
            proof_file,
            cfg("dify", "submit_workflow_app_key"),
            OPERATOR,
        )

    @staticmethod
    def build_submit_data():
        return {
            "vulnIsAffected": 0,
            "vulnWhetherToPatch": 0,
            "vulnUnaffactedCategory": "受影响代码不在执行路径",
            "vulnUnaffactedReason": UNAFFECTED_REASON,
            "vulnTriggerCondition": "",
            "vulnJudgementMethod": "重评估后无影响",
            "vulnCompileConfig": "无",
            "vulnDomain": "驱动",
            "vulnAssessor": ASSESSORS,
        }

    def submit_workflow(self, item, data, token, file_id):
        inputs = {
            "pbiId": item["pbild"],
            "vulnScoreProof": {
                "transfer_method": "local_file",
                "upload_file_id": file_id,
            },
            "vulnHwpsirtid": item["hwpsirt_id"],
            "data": json.dumps(data, ensure_ascii=False),
            "handler": OPERATOR,
            "token": token,
        }
        try:
            resp = run_workflow(
                inputs,
                cfg("dify", "submit_workflow_app_key"),
                user="dts_auto_submit",
            )
            dts_nos = self._extract_dts_nos(resp)
            if dts_nos:
                return {"status": "success", "dts_nos": dts_nos, "error": ""}
            return {
                "status": "submit_failed",
                "dts_nos": [],
                "error": self._extract_submit_error(resp),
            }
        except Exception as err:
            return {
                "status": "submit_failed",
                "dts_nos": [],
                "error": str(err),
            }

    @staticmethod
    def _extract_dts_nos(resp):
        outputs = resp.get("data", {}).get("outputs", {})
        items = outputs.get("errorlist", [])
        if not items and isinstance(outputs.get("result"), dict):
            items = outputs["result"].get("errorlist", [])
        return [
            x["dtsNo"]
            for x in items
            if isinstance(x, dict) and x.get("dtsNo")
        ]

    @staticmethod
    def _extract_submit_error(resp):
        outputs = resp.get("data", {}).get("outputs", {})
        result = outputs.get("result", {})
        sources = [result, outputs]
        sources.extend(outputs.get("errorlist", []) or [])

        for source in sources:
            if not isinstance(source, dict):
                continue
            for key in ("error", "message", "msg", "reason"):
                if source.get(key):
                    return str(source[key])
        return json.dumps(resp, ensure_ascii=False, default=str)[:1000]

    def submit_all(self, precheck):
        ready, token = precheck["ready"], precheck["token"]
        if not ready:
            return []

        self._log_step(7, "上传证明文件并执行DTS走单")
        results, file_cache, file_errors = [], {}, {}
        data = self.build_submit_data()

        for item in ready:
            proof = item["proof_name"]
            try:
                if proof in file_errors:
                    raise RuntimeError(file_errors[proof])

                if proof not in file_cache:
                    try:
                        file_cache[proof] = self.upload_proof(
                            item["proof_file"]
                        )
                        logger.info("[文件上传成功] %s", proof)
                    except Exception as err:
                        file_errors[proof] = str(err)
                        raise

                submit = self.submit_workflow(
                    item,
                    data,
                    token,
                    file_cache[proof],
                )
            except Exception as err:
                submit = {
                    "status": "upload_failed",
                    "dts_nos": [],
                    "error": str(err),
                }

            result = {**item, **submit}
            results.append(result)

            if result["status"] == "success":
                self._save_history(item, result)
                logger.info(
                    "[走单成功] cve=%s hwpsirt=%s pbild=%s dts=%s",
                    item["cve_id"],
                    item["hwpsirt_id"],
                    item["pbild"],
                    ",".join(result["dts_nos"]),
                )
            else:
                logger.error(
                    "[走单失败] cve=%s hwpsirt=%s reason=%s",
                    item["cve_id"],
                    item["hwpsirt_id"],
                    result["error"],
                )
        return results

    @staticmethod
    def _writeback_targets(results):
        targets, seen = [], set()
        for result in results:
            if result["status"] not in {"success", "already_submitted"}:
                continue
            for item in result["targets"]:
                key = (
                    item["qg_version"],
                    item["hwpsirt_id"],
                    item["cve_id"],
                )
                if key not in seen:
                    seen.add(key)
                    targets.append(item)
        return targets

    @staticmethod
    def _update_excel(local_file, items):
        wb = load_workbook(local_file)
        try:
            ws = wb["Sheet1"]
            headers = {
                str(c.value).strip(): c.column
                for c in ws[1] if c.value is not None
            }
            for col in ("HWPSIRT ID", "CVE ID", "已走单"):
                if col not in headers:
                    raise ValueError(f"{Path(local_file).name}缺少列: {col}")

            row_map = {}
            for row in range(2, ws.max_row + 1):
                hwpsirt = str(
                    ws.cell(row, headers["HWPSIRT ID"]).value or ""
                ).strip()
                cve = str(ws.cell(row, headers["CVE ID"]).value or "").strip()
                row_map.setdefault((hwpsirt, cve), []).append(row)

            missing = [
                x for x in items
                if (x["hwpsirt_id"], x["cve_id"]) not in row_map
            ]
            if missing:
                return 0, missing

            updated = 0
            for item in items:
                for row in row_map[(item["hwpsirt_id"], item["cve_id"])]:
                    cell = ws.cell(row, headers["已走单"])
                    if str(cell.value or "").strip() != "是":
                        cell.value = "是"
                        updated += 1

            if updated:
                wb.save(local_file)
            return updated, []
        finally:
            wb.close()

    def update_excel_submit_status(self, results):
        targets = self._writeback_targets(results)
        if not targets:
            return {"updated": 0, "failed": 0}

        self._log_step(8, "回写基线Excel")
        file_map = {Path(x).stem: x for x in excel_list}
        grouped = {}
        for item in targets:
            grouped.setdefault(item["qg_version"], []).append(item)

        updated_total, failed = 0, 0
        file_status = {}

        with FtpMgr.get_ias_ftp_client() as ftp:
            for qg, items in grouped.items():
                file_name = file_map.get(qg)
                if not file_name:
                    failed += 1
                    file_status[qg] = False
                    continue

                remote = str(PurePosixPath(excel_base_url) / file_name)
                try:
                    # 基于 FTP 最新内容补写，避免使用上次上传失败后已改过的本地副本。
                    local = self._download(ftp, remote, self.temp_dir.name)
                    if not local:
                        raise RuntimeError("FTP下载失败")
                    updated, missing = self._update_excel(local, items)
                    if missing:
                        raise RuntimeError(f"未找到{len(missing)}条数据")

                    if updated and not ftp.upload_file(
                            str(local),
                            remote,
                            force_upd=True,
                    ):
                        raise RuntimeError("FTP上传失败")

                    updated_total += updated
                    file_status[qg] = True
                    logger.info(
                        "[Excel回写成功] file=%s count=%s",
                        file_name,
                        updated,
                    )
                except Exception as err:
                    failed += 1
                    file_status[qg] = False
                    logger.error(
                        "[Excel回写失败] file=%s reason=%s",
                        file_name,
                        err,
                    )

        for result in results:
            if result["status"] not in {"success", "already_submitted"}:
                continue
            synced = all(
                file_status.get(x["qg_version"], False)
                for x in result["targets"]
            )
            self._set_excel_synced(result, synced)

        return {"updated": updated_total, "failed": failed}

    def send_summary(self, summary, report):
        if not summary["total"] and not summary["history_hits"]:
            return

        try:
            token = WeLinkToken.objects.get(uid='j00600960').token
            WeLinkMsg.send_msg(token, WELINK_NOTICE_NO4, report)
            logger.info("小鲁班走单结果通知发送成功")
        except Exception as err:
            logger.error("小鲁班通知失败: %s", err)

    @staticmethod
    def _log_summary(summary):
        logger.info("=" * 70)
        logger.info("DTS自动走单任务汇总")
        for key, title in (
                ("total", "本次真实走单"),
                ("success", "成功"),
                ("failed", "失败"),
                ("history_hits", "历史成功补回"),
                ("blocked", "预检查失败"),
                ("missing_files", "缺证明文件"),
                ("excel_updated", "Excel回写"),
                ("excel_failed", "Excel回写失败"),
                ("duration", "耗时"),
        ):
            logger.info("%s: %s", title, summary[key])
        logger.info("=" * 70)

    def run_auto_submit(self):
        start = time.time()
        logger.info("DTS自动走单任务开始")

        precheck = self.precheck()
        results = self.submit_all(precheck)

        writeback = results + precheck["history_hits"]
        excel = self.update_excel_submit_status(writeback)

        success = sum(
            item["status"] == "success"
            for item in results
        )
        summary = {
            "total": len(results),
            "success": success,
            "failed": len(results) - success,
            "history_hits": len(precheck["history_hits"]),
            "blocked": len(precheck["blocked"]),
            "missing_files": len(precheck["missing_files"]),
            "excel_updated": excel["updated"],
            "excel_failed": excel["failed"],
            "duration": int(time.time() - start),
        }

        report = self._build_result_report(
            precheck,
            results,
            excel,
        )

        logger.info("\n%s", report)
        self.send_summary(summary, report)

        logger.info("DTS自动走单任务结束")
        return summary
