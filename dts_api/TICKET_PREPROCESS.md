# DTS 单据预处理

`DtsAutoSubmitService.precheck()` 只对原流程筛出的候选读取 `Sheet1 → DTS单号`，使用 `ticket_preprocess.py` 查询：

```text
GET https://codeartsx.rnd.yinwang.com/DTSPortal/v1/getTicket
参数：dtsNo=<单号>，_=<当前毫秒时间戳>
```

仅当响应中的 `result.rejectFlow` 恰好包含以下两个 `eventName`，且各出现一次时命中：

- `提交开发人员定位`
- `转发测试(项目)经理审核`

事件顺序不限；多个、少于两个、重复同一事件或其他名称均不命中。
只使用 `rejectFlow.eventName` 判定，其他节点状态不作为筛选条件。

## 接入原走单流程

执行顺序：

1. 从 `Sheet1` 筛选“无影响、未走单”的记录，并保留原有“是否评审=是”和评审总表匹配条件。
2. 保留原来的合并去重、分组数据冲突及关联基线状态检查。
3. 对剩余候选中的 DTS 单号查询状态；跨行和跨文件重复的单号，本轮只查询一次。
4. 命中两个目标事件时拦截该 `pbild + HWPSIRT ID` 分组，清单状态为“待走单”。
5. 查询失败或单号异常时暂缓提交，清单状态为“待核对”。其他记录继续原有流程。

检查放在历史去重之前，命中项即使有本地成功历史也不会被补写成“已走单”。
拦截项不调用实际走单、不保存成功历史、不回写原 Excel 标记。
空白 DTS 单号沿用原流程；非空但无法识别的单号、没有缓存值的公式列为待核对。

一格多个 DTS 单号时分别查询，只要任一命中，就拦截整个提交分组。
同组存在查询失败时，失败单号仍会单独导出，避免丢失异常。

依赖安装：

```bash
python -m pip install -r dts_api/requirements-preprocess.txt
```

将更新后的 `auto_submit.py`、新增的 `ticket_preprocess.py` 放入原项目的 `test_system_third_party/service/dts_api/` 目录。
沿用原定时任务的 Django 初始化，在入口注入已登录的 Session。现有 Session 可以直接传入；
`get_dts_session(headers)` 则复用已验证的 Cookie 配置，补齐 `cftk`、AJAX 标识和 CookieJar。
凭据由运行时配置提供，不写入源码、导出清单或日志。

```python
from test_system_third_party.service.dts_api.auto_submit import DtsAutoSubmitService
from test_system_third_party.service.dts_api.ticket_preprocess import get_dts_session


def run_dts_job(dts_headers, export_dir="./dts_reports"):
    with get_dts_session(dts_headers) as session:
        with DtsAutoSubmitService(
            dts_session=session,
            export_dir=export_dir,
        ) as service:
            summary = service.run_auto_submit()

    print("待走单：", summary["pending_dts_nos"])
    print("清单路径：", summary["ticket_export_file"])
    return summary
```

首次接入时，只检查和导出、不实际提交，可把入口里的 `run_auto_submit()` 调用替换为：

```python
precheck = service.precheck()
print(precheck["ticket_preprocess"]["pending_dts_nos"])
print(precheck["ticket_export_file"])
```

存在 DTS 单号而未提供 Session 时，这些候选会暂缓提交并导出为待核对。
Session 的登录和续期由调用方负责；Windows 和 Linux 使用同一查询逻辑，网络和 CA 信任由各部署环境配置。

## 导出与汇总

默认导出到当前工作目录的 `dts_reports/`，也可通过 `export_dir` 指定可写目录。
文件名为 `dts_pending_<时间>_<唯一后缀>.xlsx`，包含：

- `待走单`：DTS 单号、状态、AOS 版本、HWPSIRT ID、CVE ID、来源行、事件名称及原因。
- `待核对`：失败或异常记录及原因，不混入待走单名单。

有效 DTS 单号去重导出；多个来源版本和行号合并展示。
导出文件不在临时目录内，服务退出后仍保留；没有待走单或待核对项时不生成空文件。
清单在获取走单 Token 和实际提交前导出；导出失败会中断，避免丢失待处理项。

`run_auto_submit()` 的汇总新增：

| 字段 | 含义 |
| --- | --- |
| `pending_dts` / `pending_dts_nos` | 去重后的待走单数量和单号列表 |
| `dts_check_failed` | 待核对条数 |
| `ticket_export_file` | 导出清单的绝对路径，没有清单时为空字符串 |

日志和原有 WeLink 汇总中增加待走单数量、单号和清单路径。仅有待走单、没有自动提交时也会发汇总。
通知最多展示前 50 个待走单号，完整名单保存在 Excel 中；版本表按各版本关联的单号统计。
清单保存在脚本运行机器，不会自动上传为 WeLink 附件。

## 单号方法与独立诊断

已有的单号方法可继续调用：

```python
from test_system_third_party.service.dts_api.ticket_preprocess import check_dts_ticket

result = check_dts_ticket("DTS0001", session=session)
# result: dts_no、query_success、matched、event_names、error
```

底层 `DtsTicketPreprocessor.run` 支持批量单号查询，查询汇总字段如下：

查询汇总字段：

| 字段 | 含义 |
| --- | --- |
| `total` | 去重后的单号数量 |
| `matched_count` / `matched_dts_nos` | 命中数量和单号列表 |
| `unmatched_count` / `unmatched_dts_nos` | 查询成功但事件不符合的数量和单号列表 |
| `failed_count` / `failures` | 查询或响应异常的数量及逐单原因 |

`run_excel` / `run_excels` 是独立的整列诊断工具，不应用走单候选条件；主流程不调用这两个方法，
而是使用原筛选结果查询，因此不会扩大本次预处理范围。

查询超时、HTTP 错误、登录跳转、非 JSON 响应或缺少必要字段会计入查询失败，并继续处理其他单号。
失败单据不会归入正常未命中列表。
缺少 `Sheet1` 或 `DTS单号` 列时直接报错，不把输入错误伪装成没有命中。

查询模块本身只读取；实际提交、原 Excel 回写和汇总通知仍由走单服务负责。
请求保留 TLS 验证以及调用环境的代理和 CA 信任配置。
真实查询需要可访问该内网域名的网络，以及有效的 DTS 登录态。

## 验证

在仓库根目录运行：

```bash
python dts_api/verify_ticket_preprocess.py
python dts_api/verify_selection.py
```

测试使用真实临时 Excel 文件、requests 适配器和本地服务替身，不访问真实 DTS、FTP 或 WeLink。
