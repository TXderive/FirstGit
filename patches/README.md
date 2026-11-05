# auto_submit.py 更新补丁

按当前代码版本选择一份补丁应用：

| 补丁 | 适用基线 | 内容 |
| --- | --- | --- |
| [auto_submit.patch](auto_submit.patch) | 本次对话最初提供的原始 `auto_submit.py` | Sheet1 优先筛选、评审记录核对、多基线双向检查，以及历史去重和 Excel 补写修复 |
| [auto_submit-history.patch](auto_submit-history.patch) | 本分支提交 `44b0d215c6a069dd1df4d4fee6588d0ea751697f` 中的 `dts_api/auto_submit.py` | 本次历史去重和 Excel 补写修复 |

两份补丁都将对应基线更新为本分支当前的 `dts_api/auto_submit.py`，只需选择一份。
本分支的完整源码已包含这些修改；补丁用于更新你自己项目中的对应文件。

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
