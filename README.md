# 周报更新工具（WeeklyReportUpdater）

这个工具从 **M&E Documents Status Tracker.xlsx** 里提取指定日期范围的提交和批复数据，自动更新技术部周例会 PPT。

## 会更新哪些内容

| 位置 | 内容 |
|---|---|
| 封面 | 汇报日期（默认取结束日期） |
| 第 2–4 页标题 | “前几周工作完成情况（M月D日 – M月D日）” |
| 1.1 采购技术文件 | MTS / MR / TBE / NMR 601 / NMR 602 两张表：Submitted 和 Approved |
| 1.2 施工图纸 | Shop Dwg 的 Submitted 和 Approved |
| 1.3 英文方案 | MS 的 Approval |
| 1.4 设计澄清 | TQ / Proposal 的 Submitted 和 Approval |
| 每页统计 | 提交、批复、通过、被拒，都按**这一页涉及的所有类型合计**；1.4 页另外统计“还未正式批复”的份数 |

第 5、6 页（下周计划、需协助事项）不改。

## 统计规则

- **提交**：`Date Submitted` 在统计范围内。
- **批复**：`Response Received Date` 在统计范围内，并且能识别出批复代码。代码优先从 `Status Latest` 读取。如果这一栏是空的（文件已经升版），就从 `JCDC Comments` 读取。
- **通过**：代码是 A / B / C / O，可以在 `config.json` 里改。表格的 Approved / Approval 栏只列通过的文件。
- **被拒**：批复数减去通过数，也就是 D / E / F / G 等。
- **还未正式批复**：`Status Latest` 是 `Open` 的行。
- **文件名**：只去掉前缀（例如 “Material Technical Submittal for”），其余文字不改。
- 没有符合条件的文件时，单元格写“无”。

## 排版

- 表格行高按内容自动调整，下方的文字框会跟着下移，避免重叠。
- 如果统计文字框被挤出页面，程序会把它挪到另一张较矮的表格下方。
- 实在放不下时，表格正文字号会逐步缩小，最小到 8 号，并在核对清单里提醒你。
- 字体、颜色、表格样式都沿用原 PPT。

## 每周怎么用

1. 双击 `WeeklyReportUpdater.exe`。
2. 选择本周的文件台账 xlsx，以及**上周生成的 PPT**（程序会自动记住上次的输出）。
3. 确认统计范围。开始日期默认是上次结束日期的第二天，结束日期默认是今天。
4. 点击“生成新 PPT”。新 PPT 和一份 `_核对清单.txt` 会保存到输出文件夹（默认 `E:\EdgeDOWNLODAD`）。
5. 看一下核对清单，再补写第 5、6 页。

### 关于每周更新的台账

- 每次点“生成”时，程序都会**重新读取**选中的 xlsx，用的是这一刻已保存的内容。在 Excel 里改完记得先保存。
- 打开程序时，会在上次台账所在的文件夹里**自动选中最新的台账**：按修改时间排序，文件名要符合 `*Status*Tracker*.xlsx`，会忽略 `~$` 开头的临时文件。所以无论是在原文件上改，还是每周另存成 `_1`、`_2`，都能选对。
- 如果选中的台账自上次生成以来没有修改过，程序会弹窗提醒，避免误用旧表。
- 日志第一行会显示这次用的是哪个文件、最后修改时间，方便确认。

核对清单里列出每个类型具体是哪几份文件、各自什么代码，也会列出不在 PPT 里、但这段时间有动态的 sheet（比如 PO、NMR 603），方便你核对。

## 修改规则（config.json）

`config.json` 和 exe 放在同一个文件夹，第一次运行时会自动生成。用记事本打开，改完保存，下次运行就生效：

- `pass_codes`：哪些代码算“通过”。
- `sections`：PPT 小节和 xlsx sheet 的对应关系。以后 PPT 里加了类型（比如 NMR 603），在这里加一行就行。
- `prefix_patterns`：要去掉的文件名前缀。
- `default_output_dir`：默认输出文件夹。
- `output_name`：输出文件名，`{date}` 会替换成封面日期。

## 命令行用法（可选）

```
WeeklyReportUpdater.exe --xlsx 台账.xlsx --ppt 上周.pptx --start 2026-10-14 --end 2026-11-03 [--date 2026-11-03] [--out 输出文件夹]
```

## 用 GitHub 打包 exe

1. 新建一个仓库，把本文件夹的所有文件上传，包括 `.github/workflows/build.yml`。
2. 推送到 `main` 分支后，GitHub Actions 会自动在 Windows 上打包。也可以在 Actions 页面手动运行“打包 exe”。
3. 打包完成后，在 Actions 这次运行的页面底部 **Artifacts** 里下载 `WeeklyReportUpdater`，里面有 exe、config.json 和 README。
4. 想要正式版本：打一个 `v` 开头的 tag（例如 `v1.0`）并推送，会自动生成 Release，附带 zip 包。

> 注意：仓库里不要上传公司的 xlsx / pptx，`.gitignore` 已经排除了这两类文件。

## 本地运行（不打包）

```
pip install -r requirements.txt
python main.py
```
