# 调研报告表格文案精简 · xlsx-report-slimming

> 把一份已经成稿的中文调研报告 Excel，改造成措辞更朴素、更像人工整理的版本。只压缩说明文字，不动数据、图片、公式和版式，并且始终另存新文件，让两版可以对照。

[![Skill](https://img.shields.io/badge/type-Agent%20Skill-blue)]()
[![Lang](https://img.shields.io/badge/lang-中文-green)]()
[![License](https://img.shields.io/badge/license-MIT-lightgrey)]()

## 这是什么

一个**流程 + 脚本型**的 Agent Skill。面向已经交付、但说明文字写得太满的中文调研报告表格，典型场景是亚马逊选品报告、竞品分析表这类多 Sheet 工作簿。

它解决的是一个很具体的问题：报告里往往既有大量数值、搜索词和嵌图，又有成段的分析文案。用常规的 Excel 库另存会丢掉 WPS 嵌图，手工改又容易漏格或串值。这套技能把「只读体检、列计划、建映射、XML 直改、全量校验、另存交付」固化成六步，并提供三个可直接运行的脚本。

## 什么时候用

- 用户附上 `.xlsx`，要求优化、精简、压缩文案、删品牌名、去掉 AI 感。
- 用户明确要求先列计划，确认后再动手。
- 用户要多个力度版本，例如先给全精简，再补一个半精简。

从零取数生成报告、改数字、增删行列或改格式，不属于本技能范围。

## 三档力度

| 用户说法 | 被改写文案保留率 | 做法 |
|---|---|---|
| 精简一些、别写这么多 | 45% 到 55% | 每格压成一句短结论，删掉推导、举例和品牌罗列 |
| 半精简、不要精简这么多 | 82% 到 90% | 只删品牌罗列、推导过程和冗余修饰，总结分析仍写足 |
| 删废话、统一措辞 | 65% 到 75% | 删铺垫句和重复论证，保留结论与数据 |

报告力度时统一用“被改写的说明文案”口径。全文件字数几乎不变是正常的，因为原始数据表和长商品 URL 占绝大部分篇幅。

## 目录结构

```
xlsx-report-slimming/
├── SKILL.md                          # 主文件，六步流程、力度对照、七条硬规则
├── README.md
├── LICENSE
├── .gitignore
├── scripts/
│   ├── inspect_workbook.py           # 步骤 1，只读体检
│   ├── rewrite_cells.py              # 步骤 4，通用 XML 直改引擎
│   └── verify_rewrite.py             # 步骤 5，十一项校验
├── references/
│   ├── editing-rules.md              # 文案口径、黑话清单、计划模板
│   └── pitfalls.md                   # 成稿改造特有的十条踩坑
└── assets/
    └── rewrite_map.example.json      # 改写映射骨架
```

## 快速上手

需要 Python 3.8 以上，只用标准库，无第三方依赖。

```bash
python scripts/inspect_workbook.py "原文件.xlsx" --out "体检报告.txt"
python scripts/rewrite_cells.py --src "原文件.xlsx" --out "原文件_半精简版.xlsx" \
    --map rewrite_map.json --backup "原目录/_backup/原文件_改前备份.xlsx"
python scripts/verify_rewrite.py --orig "原文件.xlsx" --new "原文件_半精简版.xlsx" \
    --map rewrite_map.json
```

改写映射是 JSON，键为工作表序号或表名，值为单元格坐标到新文案，骨架见 `assets/rewrite_map.example.json`。校验十一项全绿才交付。

## 几条关键约束

1. 工作簿含 `xl/cellimages.xml`、`xl/media/` 或 `xl/drawings/` 时，绝不能用 `openpyxl.save`，它会丢掉 WPS 嵌图，必须走 XML 直改。
2. 每个改写格在 sharedStrings 末尾追加独立条目并只改自己的索引，绝不复用或修改旧条目，否则所有引用同一索引的格子会一起变。
3. 原件不动，另存新文件；改前备份统一放 `_backup/` 子文件夹。
4. 所有数值、搜索词、月搜索量、购买率、SPR、价格、BSR、评分和日期保持原样与原精度。
5. 定位单元格按坐标精确匹配，不用全局单元格正则，自闭合空单元格会吞掉后一格。
6. 文件被 WPS 或 Excel 独占打开时，不要用 COM 的 `SaveAs` 指向目标路径，多窗口下对象指向不可靠。

## 依赖说明

`SKILL.md` 与 `references/pitfalls.md` 引用的 X1 到 X18 硬规则、D1 到 D7 数据纪律，来自配套的亚马逊市场调研技能，本仓库不重复维护。单独使用本仓库的三个脚本与文案口径完全可行，只是规则编号指向外部文档。

## 安装

把整个目录复制到 Agent 的 skills 目录即可，例如 `~/.workbuddy/skills/xlsx-report-slimming/`。仓库发布副本已做脱敏，示例中的项目名均替换为通用表述；品牌名保留，因为它们是校验脚本的默认黑名单值。

## License

MIT，见 [LICENSE](LICENSE)。
