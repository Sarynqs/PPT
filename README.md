# 周更

每周固定更新工作的小工具。打开后在首页选择功能：

| 功能 | 做什么 |
|---|---|
| **更新周报** | 从 M&E Documents Status Tracker 提取指定日期范围的提交 / 批复数据，更新技术部周例会 PPT |
| **更新 M&E 台账** | 用 M&E Documents Status Tracker 更新 **Transmittals List-Code G** 和 **Smart City Related Document Status** 两份表，改动处标黄 |

两个功能共用同一份 M&E 台账。程序会在上次所在的文件夹里，自动选中最新的一份。

---

## 一、更新周报

### 会更新哪些内容
- **封面**：汇报日期。
- **第 2–4 页标题**：“前几周工作完成情况（M月D日 – M月D日）”。
- **1.1 采购技术文件**：MTS / MR / TBE / NMR 601 / NMR 602 的 Submitted 和 Approved。
- **1.2 施工图纸**：Shop Dwg。
- **1.3 英文方案**：MS 的 Approval。
- **1.4 设计澄清**：TQ / Proposal。
- **每页统计**：提交、批复、通过、被拒，按这一页涉及的所有类型合计；1.4 页另外统计“还未正式批复”。
- 第 5、6 页（下周计划、需协助事项）不改。

### 统计规则
- **提交**：`Date Submitted` 在统计范围内。
- **批复**：`Response Received Date` 在统计范围内。代码优先读 `Status Latest`，空白（文件已升版）时读 `JCDC Comments`。
- **通过**：代码为 A / B / C / O，其余算被拒。Approved 栏只列通过的文件。
- **文件名**：只去掉前缀，不改其他文字。
- **排版**：行高按内容自动调整；放不下时会挪动统计框，或缩小字号（最小 8 号）。
- 每次会同时生成 `_核对清单.txt`，逐份列出文件和代码。

### 用法
选择 M&E 台账和**上周生成的 PPT**（程序会自动记住），确认日期，点“生成新 PPT”。

---

## 二、更新 M&E 台账

### 匹配方法
- 按文件编号（去掉 `/00` 这类版本后缀）在 M&E 里找同一文件的**最新版本**。
- 只有两种情况才会改：M&E 里出现了更新的 Transmittal，或者同一个 Transmittal 的状态变了。

### Transmittals List-Code G
- **只更新指定负责人、指定文件类型的行**，默认是 `陈书翰` + `SDR`。界面上可以临时改，多个用顿号隔开，留空表示全部。
- **有更新的 Transmittal 时**：
  - 「更新情况」填 `已更新`；
  - 「状态」按最新代码填：A/B/C/D/O 填 `已完成`，其余填 `已上报`；
  - 「Status」填最新代码；
  - 最后一列「Transmittal No. CHEC」填新的 Transmittal 号。
- **同一 Transmittal 只是状态变了时**：只改「Status」和「状态」。
- R 列的 Transmittal 号如果在 M&E 里对应的是别的文件，只在日志里提示“可能填错”，不会改动。
- 不会新增行，也不会改顶部的统计公式。

### Smart City Related Document Status
- **Material Submittals 页**：按 MR/MTS、TBE、AR、PO、NMR 601/602/603 各组的 Transmittal 号，在 M&E 对应的 sheet 里找到文件，再更新这一组的 Transmittal No. / Rev. / 提交日期 / 批复日期 / 状态。
- **Shopdrawing / MS / TQ / ITP 页**：按 Document No 找最新版本，更新 Transmittal、Rev.、文件编号（带 `/0x` 版本号时）、提交日期、批复日期、状态。
- 状态的写法跟原单元格一致，例如原来是 `CODE-B` 就写大写，原来是 `OPEN` 就写 `OPEN`。
- 空着的格子不会自动补：按描述匹配不够可靠。

### 标黄和更新记录
- 改动的单元格标**黄色**，每份表最后会生成「更新记录」工作表，逐格列出原值和新值。
- 勾选“清除上次更新留下的黄色标记”时，上次标的黄会恢复成原来的底色，只保留本次改动的黄色。「更新记录」里有一列隐藏的“原底色”，就是用来做这个的。
- 新文件保存到输出文件夹，文件名带当天日期，例如 `Transmittals List-Code G-20261104.xlsx`。原文件不动。
- 程序会记住这次的输出，下周打开时自动作为输入。

---

## 修改规则（config.json）

`config.json` 和 exe 放在同一个文件夹，第一次运行时会自动生成。用记事本打开，改完保存即可：

- **周报**：`pass_codes`、`sections`（小节 ↔ sheet）、`prefix_patterns`、`default_output_dir` 等。
- **M&E 台账**（`me_update` 段）：
  - `codeg_people` / `codeg_types`：默认负责人和文件类型；
  - `done_codes`：哪些代码算“已完成”；
  - `material_groups`：Material Submittals 的分组对应哪些 sheet；
  - 自动选文件用的名字规则。

## 命令行（可选）

```
周更.exe report --xlsx 台账.xlsx --ppt 上周.pptx --start 2026-10-14 --end 2026-11-03
周更.exe me --xlsx 台账.xlsx --codeg "Transmittals List-Code G.xlsx" --smartcity "Smart City ....xlsx"
```

## 用 GitHub 打包 exe

1. 把本文件夹所有文件上传到仓库，包括 `.github/workflows/build.yml`。可以直接覆盖原来的 WeeklyReportUpdater 仓库。
2. 推送后，Actions 会自动在 Windows 上打包。
3. 在 Actions 这次运行页面底部的 **Artifacts** 里下载 `ZhouGeng`，里面有 `周更.exe`、`config.json` 和 README。
4. 打 `v` 开头的 tag（例如 `v2.0`）会自动生成 Release。

> 仓库里不要上传公司的 xlsx / pptx，`.gitignore` 已经排除了。

## 本地运行

```
pip install -r requirements.txt
python main.py
```
