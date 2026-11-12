---
name: xlsx-report-slimming
description: "对已交付的中文调研报告 Excel 做文案精简，在保留全部数值、搜索词、图片、公式和版式的前提下压缩说明文字、删除品牌罗列与运营黑话，另存新版本。当用户附上 .xlsx 并说「优化这个表格」「精简一些」「来个半精简版本」「别写这么多」「删掉品牌名」「更像人工整理」或要求「先列计划，确认后再改」时使用。Slim Chinese analysis copy in an existing xlsx report while preserving all figures, embedded images and formulas, then output a separate version without overwriting the original."
agent_created: true
---

# 调研报告表格文案精简

## 用途

把一份已经成稿的中文调研报告 Excel，改造为措辞更朴素、更像人工整理的精简版本。核心是**只改说明文案，不动数据、图片、公式和版式**，并且始终另存新文件，让用户能两版对照。

## 何时使用

- 用户附上 `.xlsx` 并要求优化、精简、压缩文案、删品牌名、去掉 AI 感。
- 用户要求「先列计划，我确认后再动手」的表格改造任务。
- 用户要多个力度版本，例如先给全精简、再补一个半精简。

## 何时不用

- 从零生成或填写调研报告、需要重新取数：改用 `亚马逊表格工作流` 技能。
- 要新建一份独立工作簿、或做透视/统计/公式构建：改用对应的表格技能。
- 用户要的是改数字、增删行列、改格式：本技能只处理文案。

## 与「亚马逊表格工作流」的分工

那个技能负责「卖家精灵取数 → 生成 9 表报告」，并沉淀了 XML 直改硬规则 X1 到 X18、数据纪律 D1 到 D7、输出格式规范。**本技能直接复用这些规则，不重复维护**，尤其 X2（自闭合单元格）、X6（文件被占用）、X7（目标格预检）、X11（共享字符串破除）、X17（差集审计）。本技能补充的是成稿改造特有的流程、踩坑与校验，见 `references/pitfalls.md`。

## 流程（6 步）

### 步骤 1：只读体检，不改文件

```bash
python scripts/inspect_workbook.py "原文件.xlsx" --out "体检报告.txt"
```

通读输出，掌握：zip 条目与大小、`xl/cellimages.xml` / `xl/media/` / `xl/drawings/` 是否存在、DISPIMG 公式数、sharedStrings 的 count 与 uniqueCount、工作表名到 `sheetN.xml` 的映射、每个表的文本格数与字数、长文本清单、品牌词与黑话与禁用符号残留。

体检结果决定技术路线：存在嵌图部件就必须 XML 直改，绝不能用 openpyxl 保存。

### 步骤 2：列计划并等待确认（硬门槛）

用户说「先列计划」「确认后再改」时，**到此为止，不得写文件**。计划至少包含六节：文件现状、精简口径、逐表改动范围、技术路线、交付与校验、需要确认的点。确认项要给出默认建议，让用户能直接回一句「按默认来」。模板见 `references/editing-rules.md`。

### 步骤 3：建立改写映射

按体检结果写 `rewrite_map.json`，键为工作表序号或表名，值为坐标到新文案。骨架见 `assets/rewrite_map.example.json`。

只列真正要改的格。表头、章节标题、原始数据表、商品 URL、listing 标题、数值、公式、图片、合并区一律不进映射。文案口径细则见 `references/editing-rules.md`。

### 步骤 4：执行改写

```bash
python scripts/rewrite_cells.py --src "原文件.xlsx" --out "原文件_半精简版.xlsx" \
    --map rewrite_map.json --backup "原目录/_backup/原文件_改前备份.xlsx"
```

引擎内置目标格预检、按坐标精确定位、每格追加独立 `<si>`、`count` 与 `uniqueCount` 重算、其余 zip 条目字节原样复制。预检不通过会直接报错退出，不会写出半成品。

### 步骤 5：校验，全绿才交付

```bash
python scripts/verify_rewrite.py --orig "原文件.xlsx" --new "原文件_半精简版.xlsx" --map rewrite_map.json
```

校验项：zip 完整性、条目名一致、图片与嵌图部件字节零变化、公式数、sharedStrings 属性、未改写工作表字节一致、目标格值精确匹配、非目标文本格值不变、字数比例、品牌与黑话与符号残留、openpyxl 只读可打开。任一项失败，回到步骤 4 修复，不要带病交付。

### 步骤 6：交付

用 `present_files` 打开新文件。回复里交代：改了几格、覆盖哪几表、力度、保留了什么、品牌与黑话残留为何仍有、图片公式完整性、备份位置、原件未改动。

## 力度选择

| 用户说法 | 被改写文案保留率 | 做法 |
|---|---|---|
| 精简一些、别写这么多 | 45% 到 55% | 每格压成一句短结论，删掉全部推导、举例和品牌罗列 |
| 半精简、不要精简这么多 | 82% 到 90% | 只删品牌罗列、推导过程、冗余修饰；总结分析仍写足 |
| 删废话、统一措辞 | 65% 到 75% | 删铺垫句和重复论证，保留结论与数据 |

**报告比例时统一用「被改写的说明文案」口径**。全文件字数通常几乎不变，因为原始数据表和长 URL 占绝大部分篇幅，用全文件字数会让效果看起来像没改。

## 硬规则

1. 有 `xl/cellimages.xml`、`xl/media/`、`xl/drawings/` 任一存在时，绝不用 `openpyxl.save`；它会丢掉 WPS 嵌图。先体检再决定路线。
2. 改写前必须做目标格存在性预检，缺失即停止（X7）。定位单元格用按坐标精确 `re.search`，不要用全局 `<c ...>(.*?)</c>` 替换，自闭合单元格会吞掉后一格（见 `references/pitfalls.md`）。
3. 每个改写格追加独立 `<si>` 并只改该格 `<v>`，绝不复用或修改旧 `<si>`（X11）。`count` 保持不变，`uniqueCount` 等于原值加改写格数。
4. 原件不动，另存新文件；改前备份统一放 `_backup/` 子文件夹，不在主目录散落。
5. 数据不足留空，不推断、不编造（D1 到 D3）。所有数值、搜索词、月搜索量、购买率、需供比、SPR、价格、销量、BSR、评分、日期保持原样与原精度。
6. 交付前扫描禁用符号、运营黑话、品牌残留。合规检索清单、listing 标题、商品 URL 里的品牌名是有意保留，交付时要说明，不要当成漏改。
7. 遇到文件被 WPS 或 Excel 独占打开，按 X6 的降级链处理；**不要用 COM 的 `ActiveWorkbook.SaveAs` 指向目标路径**，多窗口下对象指向不可靠，会把别的工作簿覆盖过去。

## 文件结构

```
xlsx-report-slimming/
├── SKILL.md                          # 本文件，主流程
├── scripts/
│   ├── inspect_workbook.py           # 步骤1 只读体检
│   ├── rewrite_cells.py              # 步骤4 XML 直改引擎
│   └── verify_rewrite.py             # 步骤5 全量校验
├── references/
│   ├── editing-rules.md              # 文案口径、力度、计划模板、格式偏好
│   └── pitfalls.md                   # 成稿改造特有的踩坑与规避
└── assets/
    └── rewrite_map.example.json      # 改写映射骨架
```

按需加载 references，不必每次全读。**但 SKILL.md 的「硬规则」和「力度选择」每次必读。**

## 修订记录

- 初版。沉淀自两次实战：某节日袜报告的全精简与半精简（10,582 字与 17,685 字两版对照），以及某宠物服饰报告的半精简（59 格、被改写文案保留 89.1%、81 张图片与 38 个 DISPIMG 公式字节零变化）。三个脚本已在真实文件上跑通，校验 11 项全绿。
