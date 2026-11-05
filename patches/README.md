# auto_submit.py 更新补丁

## 预处理拦截与待走单导出集成版

[dts-preprocess.patch](dts-preprocess.patch) 基于提交 `b80d5a5` 中的 `dts_api/auto_submit.py`。
它更新走单流程，并新增 `ticket_preprocess.py` 和 `requirements-preprocess.txt`。
在原项目 `auto_submit.py` 所在目录执行：

```bash
git apply --check dts-preprocess.patch
git apply dts-preprocess.patch
python -m pip install -r requirements-preprocess.txt
```

随后在现有定时任务入口注入已登录 Session 和导出目录，调用示例见 [预处理集成说明](../dts_api/TICKET_PREPROCESS.md)。

## 之前的历史去重补丁

按当前代码版本选择一份补丁应用：

| 补丁 | 适用基线 | 内容 |
| --- | --- | --- |
| [auto_submit.patch](auto_submit.patch) | 本次对话最初提供的原始 `auto_submit.py` | Sheet1 优先筛选、评审记录核对、多基线双向检查，以及历史去重和 Excel 补写修复 |
| [auto_submit-history.patch](auto_submit-history.patch) | 本分支提交 `44b0d215c6a069dd1df4d4fee6588d0ea751697f` 中的 `dts_api/auto_submit.py` | 本次历史去重和 Excel 补写修复 |

这两份历史补丁分别将对应基线更新为提交 `b80d5a5` 中的 `dts_api/auto_submit.py`，只需选择一份。
当前分支源码在此基础上继续包含预处理集成；这些补丁用于按版本更新你自己的项目。

## 应用方法

下载所选补丁，放到自己项目的 `auto_submit.py` 所在目录，在该目录执行。
以下以完整补丁为例；选择增量补丁时，将文件名替换为 `auto_submit-history.patch`：

```bash
git apply --check auto_submit.patch
git apply auto_submit.patch
```

两份补丁均已在对应基线的临时 Git 仓库中验证可检查、可应用，应用结果与完整源码一致。

## 本次防重复走单修复

- 以 `pbild + HWPSIRT ID` 查询 MongoDB 成功历史，命中后只补写当前 Excel 目标，不再次提交 DTS。
- `excel_synced=False` 不再成为重新走单的理由；Excel 标记被覆盖后也可补写。
- Excel 中匹配同一 `HWPSIRT ID + CVE ID` 的重复行全部标记为已走单。
- 回写前重新下载 FTP 最新文件，避免使用上传失败后已修改的本地副本。

42 项本地测试通过，包括连续多天 FTP 回写失败期间不重复提交 DTS 的模拟流程。
该保护依赖已保存的成功历史；未入库的旧单或工作流超时后的实际 DTS 状态需要单独核对。
