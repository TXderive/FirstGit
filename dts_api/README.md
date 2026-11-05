# DTS 自动走单

`auto_submit.py` 是修改后的完整 `DtsAutoSubmitService`，供现有 Django 项目集成使用。
运行时依赖原项目的 MongoDB、FTP、Dify、WeLink 及相关配置模块。

筛选与核对顺序：

1. 从 `Sheet1` 筛选 `最新状态=无影响` 且 `已走单` 为空的记录。
2. 按 `HWPSIRT ID + CVE ID` 到 `评审总表` 核对对应记录。
3. 从评审总表的 `原始CVSS评分` 列读取会议纪要文件名，保留原有 `Sheet1` 的 `是否评审=是` 条件。
4. 会议纪要下载和其余预检查通过后，进入原有上传、走单、历史保存、Excel 回写及通知流程。

## 本地验证

验证脚本需要 `pandas` 和 `openpyxl`，使用临时 Excel 文件检查筛选规则及后续分组、预检查行为。
它仅加载数据处理方法，不会执行实际走单或发送通知。

在仓库根目录运行：

```bash
python dts_api/verify_selection.py
```
