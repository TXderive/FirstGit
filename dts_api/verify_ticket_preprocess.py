"""用真实 Excel 和 requests 测试适配器验证预处理，不访问 DTS 内网。"""

import json
import os
import tempfile
import unittest
from http.client import HTTPMessage
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch
from urllib.parse import parse_qs, urlsplit

import pandas as pd
import requests
from requests.adapters import BaseAdapter

from ticket_preprocess import (
    DtsTicketPreprocessor,
    DtsTicketQueryError,
    GET_TICKET_URL,
    check_dts_ticket,
    get_dts_session,
    matches_target_events,
    read_excel_dts_numbers,
)


EVENT_A = "提交开发人员定位"
EVENT_B = "转发测试(项目)经理审核"


def ticket(events=(EVENT_A, EVENT_B)):
    return {"result": {
        "activeNodeList": [{"nodeName": "问题提交人填写"}],
        "dtsStatus": {"nodeName": "测试（项目）经理审核"},
        "rejectFlow": [{"eventName": name} for name in events],
    }}


class FixtureAdapter(BaseAdapter):
    def __init__(self, replies):
        self.replies = replies
        self.calls = []
        self.closed = False

    def send(self, request, **kwargs):
        self.calls.append((request, kwargs))
        number = parse_qs(urlsplit(request.url).query)["dtsNo"][0]
        reply = self.replies[number]
        if isinstance(reply, Exception):
            raise reply
        if isinstance(reply, tuple):
            status, body, headers = reply
        else:
            status, body, headers = 200, reply, {}
        response = requests.Response()
        response.status_code = status
        response.url = request.url
        response.request = request
        response.headers.update(headers)
        message = HTTPMessage()
        for name, value in headers.items():
            message.add_header(name, value)
        response.raw = SimpleNamespace(_original_response=SimpleNamespace(msg=message))
        response._content = body if isinstance(body, bytes) else json.dumps(body).encode("utf-8")
        response._content_consumed = True
        return response

    def close(self):
        self.closed = True


class EventMatchingTests(unittest.TestCase):
    def test_exact_two_event_names_match(self):
        self.assertTrue(matches_target_events(ticket()))

    def test_order_does_not_matter(self):
        self.assertTrue(matches_target_events(ticket((EVENT_B, EVENT_A))))

    def test_extra_event_does_not_match(self):
        self.assertFalse(matches_target_events(ticket((EVENT_A, EVENT_B, "其他操作"))))

    def test_extra_duplicate_event_does_not_match(self):
        self.assertFalse(matches_target_events(ticket((EVENT_A, EVENT_B, EVENT_A))))

    def test_missing_or_repeated_event_does_not_match(self):
        for events in ((), (EVENT_A,), (EVENT_A, EVENT_A), (EVENT_A, "其他操作")):
            with self.subTest(events=events):
                self.assertFalse(matches_target_events(ticket(events)))

    def test_surrounding_whitespace_is_trimmed(self):
        self.assertTrue(matches_target_events(ticket((f" {EVENT_A} ", f"\n{EVENT_B}"))))

    def test_other_status_fields_do_not_control_matching(self):
        payload = ticket()
        payload["result"]["activeNodeList"] = []
        payload["result"]["dtsStatus"] = {"nodeName": "其他节点"}
        self.assertTrue(matches_target_events(payload))

    def test_malformed_response_is_an_error(self):
        cases = (
            None, [], {}, {"result": None}, {"result": {}},
            {"result": {"rejectFlow": None}},
            {"result": {"rejectFlow": {"eventName": EVENT_A}}},
            {"result": {"rejectFlow": [None]}},
            {"result": {"rejectFlow": [{"eventName": 1}]}},
            {"result": {"rejectFlow": [{"eventName": " "}]}},
        )
        for payload in cases:
            with self.subTest(payload=payload):
                with self.assertRaises(DtsTicketQueryError):
                    matches_target_events(payload)


class QueryTests(unittest.TestCase):
    def client(self, replies, **kwargs):
        session = requests.Session()
        adapter = FixtureAdapter(replies)
        session.mount("https://", adapter)
        self.addCleanup(session.close)
        return DtsTicketPreprocessor(session=session, **kwargs), adapter

    def test_request_uses_ticket_param_fresh_timestamp_timeout_and_tls(self):
        client, adapter = self.client({"DTS0001": ticket()})
        with patch("ticket_preprocess.time.time", return_value=1793862861.147):
            result = client.run(["DTS0001"])
        self.assertEqual(result["matched_dts_nos"], ["DTS0001"])
        request, options = adapter.calls[0]
        self.assertEqual(request.method, "GET")
        self.assertEqual(request.url.split("?")[0], GET_TICKET_URL)
        self.assertEqual(parse_qs(urlsplit(request.url).query), {
            "dtsNo": ["DTS0001"], "_": ["1793862861147"],
        })
        self.assertEqual(options["timeout"], (5, 30))
        self.assertEqual(
            options["verify"],
            os.environ.get("REQUESTS_CA_BUNDLE") or os.environ.get("CURL_CA_BUNDLE") or True,
        )

    def test_session_authentication_is_reused_without_mutation(self):
        client, adapter = self.client({"DTS0001": ticket()}, headers={"X-Extra": "test-only"})
        client.session.headers["X-Test-Auth"] = "test-only"
        client.session.cookies.set("session", "test-only", domain="codeartsx.rnd.yinwang.com")
        client.run(["DTS0001"])
        headers = adapter.calls[0][0].headers
        self.assertEqual(headers["X-Test-Auth"], "test-only")
        self.assertEqual(headers["X-Extra"], "test-only")
        self.assertIn("session=test-only", headers["Cookie"])
        self.assertNotIn("X-Extra", client.session.headers)

    def test_batch_separates_matches_nonmatches_and_errors(self):
        client, adapter = self.client({
            "DTS0001": ticket(), "DTS0002": ticket((EVENT_A,)),
            "DTS0003": requests.Timeout("test timeout"),
            "DTS0004": ticket((EVENT_B, EVENT_A)),
        })
        result = client.run(["DTS0001", "DTS0002", "DTS0003", "DTS0004"])
        self.assertEqual(result["total"], 4)
        self.assertEqual(result["matched_count"], 2)
        self.assertEqual(result["matched_dts_nos"], ["DTS0001", "DTS0004"])
        self.assertEqual(result["unmatched_dts_nos"], ["DTS0002"])
        self.assertEqual(result["failed_count"], 1)
        self.assertEqual(result["failures"][0]["dts_no"], "DTS0003")
        self.assertEqual(len(adapter.calls), 4)

    def test_duplicates_and_empty_values_are_not_queried_twice(self):
        client, adapter = self.client({"DTS0001": ticket()})
        result = client.run([" DTS0001 ", None, "", "dts0001"])
        self.assertEqual(result["total"], 1)
        self.assertEqual(len(adapter.calls), 1)

    def test_single_string_is_one_ticket(self):
        client, adapter = self.client({"DTS0001": ticket()})
        self.assertEqual(client.run("DTS0001")["matched_count"], 1)
        self.assertEqual(len(adapter.calls), 1)

    def test_invalid_input_is_reported_without_request(self):
        client, adapter = self.client({})
        result = client.run(["not-a-ticket"])
        self.assertEqual(result["failed_count"], 1)
        self.assertEqual(adapter.calls, [])

    def test_empty_batch_does_not_send_requests(self):
        client, adapter = self.client({})
        result = client.run([])
        self.assertEqual(result["total"], 0)
        self.assertEqual(adapter.calls, [])

    def test_http_errors_and_login_redirects_are_failures(self):
        for status in (302, 401, 403, 500):
            with self.subTest(status=status):
                client, adapter = self.client({
                    "DTS0001": (status, b"login", {"Location": "https://login.invalid/"}),
                })
                result = client.run(["DTS0001"])
                self.assertEqual(result["failed_count"], 1)
                self.assertEqual(result["unmatched_count"], 0)
                self.assertIn(f"HTTP {status}", result["failures"][0]["error"])
                self.assertEqual(len(adapter.calls), 1)

    def test_html_login_page_is_not_treated_as_nonmatch(self):
        client, _ = self.client({"DTS0001": (200, b"<html>Login</html>", {})})
        result = client.run(["DTS0001"])
        self.assertEqual(result["failed_count"], 1)
        self.assertEqual(result["unmatched_count"], 0)

    def test_invalid_business_response_is_reported_as_failure(self):
        client, _ = self.client({"DTS0001": {"result": {}}})
        result = client.run(["DTS0001"])
        self.assertEqual(result["failed_count"], 1)
        self.assertEqual(result["unmatched_count"], 0)

    def test_borrowed_session_is_not_closed(self):
        client, adapter = self.client({"DTS0001": ticket()})
        with client:
            client.run(["DTS0001"])
        self.assertFalse(adapter.closed)

    def test_owned_session_is_closed(self):
        session = Mock()
        with patch("ticket_preprocess.requests.Session", return_value=session):
            with DtsTicketPreprocessor():
                pass
        session.close.assert_called_once_with()

    def test_local_single_ticket_function_contract(self):
        client, _ = self.client({"DTS0001": ticket()})
        result = check_dts_ticket("DTS0001", session=client.session)
        self.assertTrue(result["query_success"])
        self.assertTrue(result["matched"])
        self.assertEqual(result["event_names"], [EVENT_A, EVENT_B])
        self.assertEqual(result["error"], "")

    def test_cookie_headers_supply_cftk_ajax_and_accept_updated_jsessionid(self):
        headers = {"Cookie": "prod_cftk=test-left#test-right; JSESSIONID=test-old"}
        adapter = FixtureAdapter({
            "DTS0001": (200, ticket(), {"Set-Cookie": "JSESSIONID=test-new; Path=/; Secure; HttpOnly"}),
        })
        with DtsTicketPreprocessor(headers=headers) as client:
            client.session.mount("https://", adapter)
            self.assertTrue(client.check("DTS0001")["matched"])
            self.assertTrue(client.check("DTS0001")["matched"])
            self.assertNotIn("Cookie", client.session.headers)
        first = adapter.calls[0][0].headers
        second = adapter.calls[1][0].headers
        self.assertEqual(first["cftk"], "test-left#test-right")
        self.assertEqual(first["X-Requested-With"], "XMLHttpRequest")
        self.assertEqual(first["Content-Type"], "application/json; charset=UTF-8")
        self.assertIn("JSESSIONID=test-old", first["Cookie"])
        self.assertIn("JSESSIONID=test-new", second["Cookie"])
        self.assertNotIn("test-old", second["Cookie"])
        self.assertEqual(headers["Cookie"], "prod_cftk=test-left#test-right; JSESSIONID=test-old")

    def test_session_builder_reports_missing_cookie_or_cftk(self):
        for headers in ({}, {"Cookie": "JSESSIONID=test-only"}, {"Cookie": "not-a-cookie"}):
            with self.subTest(header_names=list(headers)):
                with self.assertRaises(ValueError):
                    get_dts_session(headers)


class ExcelTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "baseline.xlsx"

    def workbook(self, values, header="DTS单号", sheet_name="Sheet1"):
        pd.DataFrame({header: values, "备注": ["keep row"] * len(values)}).to_excel(
            self.path, sheet_name=sheet_name, index=False,
        )

    def test_column_extraction_splits_cells_deduplicates_and_reports_invalid_rows(self):
        self.workbook([" DTS0001 ", None, "DTS0002，dts0001;DTS0003\nDTS0004", "暂无", "DTS0001abc"])
        source = read_excel_dts_numbers(self.path)
        self.assertEqual(source["dts_nos"], ["DTS0001", "DTS0002", "DTS0003", "DTS0004"])
        self.assertEqual(source["empty_cells"], 1)
        self.assertEqual(source["invalid_cells"], [
            {"row": 5, "value": "暂无"}, {"row": 6, "value": "DTS0001abc"},
        ])

    def test_header_whitespace_is_trimmed(self):
        self.workbook(["DTS0001"], header=" DTS单号 ")
        self.assertEqual(read_excel_dts_numbers(self.path)["dts_nos"], ["DTS0001"])

    def test_missing_ticket_column_is_an_error(self):
        self.workbook(["DTS0001"], header="其他列")
        with self.assertRaisesRegex(ValueError, "缺少 DTS单号"):
            read_excel_dts_numbers(self.path)

    def test_missing_sheet1_is_an_error(self):
        self.workbook(["DTS0001"], sheet_name="Sheet2")
        with self.assertRaises(ValueError):
            read_excel_dts_numbers(self.path)

    def test_excel_to_query_summary_end_to_end(self):
        self.workbook(["DTS0001", "DTS0002", "dts0001", "暂无"])
        session = requests.Session()
        self.addCleanup(session.close)
        adapter = FixtureAdapter({"DTS0001": ticket(), "DTS0002": ticket((EVENT_A,))})
        session.mount("https://", adapter)
        with DtsTicketPreprocessor(session=session) as client:
            result = client.run_excel(self.path)
        self.assertEqual(result["matched_dts_nos"], ["DTS0001"])
        self.assertEqual(result["unmatched_dts_nos"], ["DTS0002"])
        self.assertEqual(result["invalid_cells"], [{"row": 5, "value": "暂无"}])
        self.assertEqual(len(adapter.calls), 2)

    def test_multiple_excels_deduplicate_queries_across_files(self):
        self.workbook(["DTS0001", "DTS0002"])
        second = Path(self.directory.name) / "second.xlsx"
        pd.DataFrame({"DTS单号": ["DTS0002", "DTS0003"]}).to_excel(
            second, sheet_name="Sheet1", index=False,
        )
        session = requests.Session()
        self.addCleanup(session.close)
        adapter = FixtureAdapter({
            "DTS0001": ticket(), "DTS0002": ticket((EVENT_A,)), "DTS0003": ticket(),
        })
        session.mount("https://", adapter)
        result = DtsTicketPreprocessor(session=session).run_excels([self.path, second])
        self.assertEqual(result["total"], 3)
        self.assertEqual(result["matched_dts_nos"], ["DTS0001", "DTS0003"])
        self.assertEqual(len(result["sources"]), 2)
        self.assertEqual(len(adapter.calls), 3)

    def test_bad_excel_fails_before_any_batch_query(self):
        self.workbook(["DTS0001"], header="其他列")
        session = requests.Session()
        self.addCleanup(session.close)
        adapter = FixtureAdapter({})
        session.mount("https://", adapter)
        with self.assertRaises(ValueError):
            DtsTicketPreprocessor(session=session).run_excels([self.path])
        self.assertEqual(adapter.calls, [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
