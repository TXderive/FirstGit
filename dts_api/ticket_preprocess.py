"""从基线 Excel 收集 DTS 单号，按 getTicket 返回的可操作事件筛选。"""

import re
import time
from http.cookies import SimpleCookie
from pathlib import Path

import pandas as pd
import requests
from requests.structures import CaseInsensitiveDict


GET_TICKET_URL = "https://codeartsx.rnd.yinwang.com/DTSPortal/v1/getTicket"
TARGET_EVENT_NAMES = frozenset({
    "提交开发人员定位",
    "转发测试(项目)经理审核",
})
DTS_NO_PATTERN = re.compile(r"(?<![A-Za-z0-9_])DTS[0-9]+(?![A-Za-z0-9_])", re.IGNORECASE)


def extract_dts_numbers(value):
    """从单元格提取单号并去重，保持首次出现的顺序。"""
    return list(dict.fromkeys(number.upper() for number in DTS_NO_PATTERN.findall(str(value))))


def get_dts_session(headers):
    """从本地配置构造带 cftk、AJAX 标识和可更新 CookieJar 的 Session。"""
    provided = CaseInsensitiveDict(headers)
    raw_cookie = provided.pop("Cookie", "")
    if not isinstance(raw_cookie, str) or not raw_cookie.strip():
        raise ValueError("headers 中需要填写本地最新的 Cookie")
    cookies = SimpleCookie()
    cookies.load(raw_cookie)
    if not cookies:
        raise ValueError("Cookie 解析为空，请只填写 Cookie 请求头的值")
    cftk = provided.get("cftk")
    if not cftk and "prod_cftk" in cookies:
        cftk = cookies["prod_cftk"].value
    if not cftk:
        raise ValueError("缺少同一次 getTicket 请求的 prod_cftk Cookie 或 cftk 请求头")
    for name in ("Host", "Connection", "Content-Length", "Accept-Encoding"):
        provided.pop(name, None)
    session = requests.Session()
    session.headers.update({
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "X-Language": "zh-CN",
        "Referer": "https://codeartsx.rnd.yinwang.com/DTSPortal/",
    })
    session.headers.update(provided)
    session.headers["cftk"] = cftk
    session.headers["Accept-Encoding"] = "gzip, deflate"
    for name, morsel in cookies.items():
        session.cookies.set(
            name, morsel.value, domain="codeartsx.rnd.yinwang.com", path="/", secure=True,
        )
    return session


class DtsTicketQueryError(Exception):
    """DTS 请求失败或返回的业务结构无法判定。"""


def matches_target_events(payload):
    """rejectFlow 必须恰好包含两个目标事件，各一次，顺序不限。"""
    if not isinstance(payload, dict) or not isinstance(payload.get("result"), dict):
        raise DtsTicketQueryError("响应缺少有效的 result 对象")
    flows = payload["result"].get("rejectFlow")
    if not isinstance(flows, list):
        raise DtsTicketQueryError("响应缺少有效的 result.rejectFlow 列表")

    names = []
    for index, flow in enumerate(flows):
        if not isinstance(flow, dict) or not isinstance(flow.get("eventName"), str):
            raise DtsTicketQueryError(f"rejectFlow[{index}] 缺少有效的 eventName")
        name = flow["eventName"].strip()
        if not name:
            raise DtsTicketQueryError(f"rejectFlow[{index}].eventName 为空")
        names.append(name)
    return len(names) == 2 and frozenset(names) == TARGET_EVENT_NAMES


def read_excel_dts_numbers(excel_file):
    """扫描 Sheet1 的 DTS单号 列；空白跳过，多单号拆分，并保留无效行位置。"""
    data = pd.read_excel(
        excel_file, sheet_name="Sheet1", dtype=str, keep_default_na=False,
    )
    data.columns = [str(column).strip() for column in data.columns]
    if "DTS单号" not in data.columns:
        raise ValueError(f"{Path(excel_file).name} 的 Sheet1 缺少 DTS单号 列")

    numbers, seen, invalid_cells = [], set(), []
    empty_cells = 0
    for row_number, value in enumerate(data["DTS单号"], start=2):
        text = str(value).strip()
        if not text:
            empty_cells += 1
            continue
        matches = extract_dts_numbers(text)
        if not matches:
            invalid_cells.append({"row": row_number, "value": text})
            continue
        for number in matches:
            number = number.upper()
            if number not in seen:
                seen.add(number)
                numbers.append(number)
    return {
        "dts_nos": numbers,
        "empty_cells": empty_cells,
        "invalid_cells": invalid_cells,
    }


class DtsTicketPreprocessor:
    """顺序查询单号并汇总；认证通过调用方现有的 Session 或 headers 复用。"""

    def __init__(self, session=None, headers=None, timeout=(5, 30), endpoint=GET_TICKET_URL):
        cookie_headers = session is None and headers and CaseInsensitiveDict(headers).get("Cookie")
        if cookie_headers:
            self.session = get_dts_session(headers)
            self.headers = {}
        else:
            self.session = session if session is not None else requests.Session()
            self.headers = dict(headers) if headers is not None else {}
        self.timeout = timeout
        self.endpoint = endpoint
        self._owns_session = session is None

    def __enter__(self):
        return self

    def __exit__(self, *_):
        if self._owns_session:
            self.session.close()

    def fetch_ticket(self, dts_no):
        try:
            with self.session.get(
                self.endpoint,
                params={"dtsNo": dts_no, "_": int(time.time() * 1000)},
                headers=self.headers,
                timeout=self.timeout,
                allow_redirects=False,
            ) as response:
                if not 200 <= response.status_code < 300:
                    raise DtsTicketQueryError(f"getTicket 返回 HTTP {response.status_code}")
                try:
                    return response.json()
                except ValueError as error:
                    raise DtsTicketQueryError("getTicket 响应不是有效 JSON，请检查登录态") from error
        except requests.RequestException as error:
            # 不把认证头、Cookie 或完整响应写入错误列表。
            raise DtsTicketQueryError(f"getTicket 请求失败: {type(error).__name__}") from error

    def check(self, dts_no):
        number = str(dts_no).strip().upper()
        result = {
            "dts_no": number, "query_success": False, "matched": False,
            "event_names": [], "error": "",
        }
        if DTS_NO_PATTERN.fullmatch(number) is None:
            result["error"] = "无效的 DTS 单号"
            return result
        try:
            payload = self.fetch_ticket(number)
            result["matched"] = matches_target_events(payload)
            result["event_names"] = [
                item["eventName"].strip() for item in payload["result"]["rejectFlow"]
            ]
            result["query_success"] = True
        except DtsTicketQueryError as error:
            result["error"] = str(error)
        return result

    def run(self, dts_nos):
        if isinstance(dts_nos, str):
            dts_nos = [dts_nos]
        numbers, seen = [], set()
        for value in dts_nos:
            number = str(value).strip().upper() if value is not None else ""
            if number and number not in seen:
                numbers.append(number)
                seen.add(number)

        matched, unmatched, failures = [], [], []
        for number in numbers:
            result = self.check(number)
            if not result["query_success"]:
                failures.append({"dts_no": number, "error": result["error"]})
                continue
            (matched if result["matched"] else unmatched).append(number)

        return {
            "total": len(numbers),
            "matched_count": len(matched),
            "matched_dts_nos": matched,
            "unmatched_count": len(unmatched),
            "unmatched_dts_nos": unmatched,
            "failed_count": len(failures),
            "failures": failures,
        }

    def run_excel(self, excel_file):
        source = read_excel_dts_numbers(excel_file)
        summary = self.run(source["dts_nos"])
        summary.update({
            "source_file": str(excel_file),
            "empty_cells": source["empty_cells"],
            "invalid_cells": source["invalid_cells"],
        })
        return summary

    def run_excels(self, excel_files):
        """合并多个已下载的 Excel，跨文件的同一单号只查询一次。"""
        if isinstance(excel_files, (str, Path)):
            excel_files = [excel_files]
        numbers, sources = [], []
        for excel_file in excel_files:
            source = read_excel_dts_numbers(excel_file)
            numbers.extend(source["dts_nos"])
            sources.append({
                "source_file": str(excel_file),
                "empty_cells": source["empty_cells"],
                "invalid_cells": source["invalid_cells"],
            })
        summary = self.run(numbers)
        summary["sources"] = sources
        return summary


def check_dts_ticket(dts_no, headers=None, session=None, timeout=(5, 30)):
    """与已在本地验证的单号查询方法保持相同的调用和返回结构。"""
    with DtsTicketPreprocessor(session=session, headers=headers, timeout=timeout) as processor:
        return processor.check(dts_no)
