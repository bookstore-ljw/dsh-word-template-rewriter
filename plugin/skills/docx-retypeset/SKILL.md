---
name: docx-retypeset
description: 把一篇 Word 论文/报告整体换到另一套模板格式（换期刊模板、单双栏互转、换字号与编号体系）：模板只提供样式与页面设置，正文全部来自资料源，公式/图/表/参考文献原样保留。通过 dsh_word retypeset 执行，并用真实 Word 复核分页与打印结果。
whenToUse: 用户给「一份原文 + 一份模板」并要求"按模板格式重写一份""换成 XX 期刊格式""双栏改单栏""标题作者栏保留模板占位"时使用。若只是把内容填进模板已有段落槽位，用 word-pipeline 技能的 fill 流程。
---

# 按模板重排整篇文档（retypeset）

`fill` 是"往格子里填内容"，`retypeset` 是"整篇换版式"：**模板贡献样式表、页面设置、
编号体系和题注格式；资料贡献每一个标题、段落、公式、图、表**。run 级格式
（斜体、上下标、域、OMML 公式）按原样搬运，不重写一个字。

## 一句话用法

```powershell
python -m dsh_word retypeset `
  --template <模板.docx> --source <原文.docx> --out <成稿.docx> `
  --profile ieee-to-ttp --verify --pdf <预览.pdf> --json <报告.json>
```

`--verify` 会用真实 Word 复核：页数、每张表是否跨页、每张图与图题是否同页，
并导出 PDF 扫描是否残留编辑标记；`issues` 非空时退出码为 1。

## 先看清两件事

1. `python -m dsh_word analyze --docx <原文>` 看资料用了哪些**样式名**（`paper title`、
   `Author`、`Abstract`、`figure caption`、`equation`、`Bibliography`…）。profile 靠样式名
   识别区块，不靠段落序号。
2. `analyze --docx <模板>` 看模板有哪些**目标样式名**（`TTP Title`、`TTP Abstract`、
   `Heading 1`、`TTP Paragraph (others)`、`Reference`…）。两者对不上就改 profile。

## profile 是什么

`dsh_word/profiles/*.json`，描述"资料样式 → 模板样式"的对应关系。新建一套格式只需
复制一份改名字，代码不用动。内置 `ieee-to-ttp.json`（IEEE 双栏 → Trans Tech 单栏）。

```jsonc
{
  "target": {                        // 模板侧：按样式名找段落原型
    "title": "TTP Title",
    "author": "TTP Author(s)",
    "address": "TTP Address",
    "abstract": "TTP Abstract",
    "keywords": "TTP Keywords",
    "heading": {"1": "Heading 1", "2": "Heading 2", "3": "Heading 3"},
    "body": "TTP Paragraph (others)",
    "caption": "TTP Paragraph (others)",
    "equation": "TTP Paragraph (others)",
    "reference": "Reference",
    "reference_heading": "Heading 1"
  },
  "source": {                        // 资料侧：样式名（不区分大小写）
    "title_styles": ["paper title", "Title"],
    "author_style": "Author",
    "abstract_styles": ["Abstract"],
    "keywords_styles": ["Keywords"],
    "heading_prefix": "Heading",     // 层级取样式名里的数字
    "figure_caption_styles": ["figure caption"],
    "table_caption_styles": ["table head"],
    "equation_styles": ["equation"],
    "reference_styles": ["Bibliography", "references", "Reference"],
    "reference_heading_text": "References",
    "drop_styles": ["sponsors"],
    "drop_sections": ["background"],
    "header_skip_patterns": ["contributed equally", "作者名"],
    "abstract_lead": "Abstract. ",
    "keywords_lead": "Keywords:"
  },
  "layout": {"figure_width_cm": 12.0, "table_width_twip": 6740, "table_font_pt": 10},
  "flags": {
    "keep_next": true,               // 表格/图题「与下段同页」
    "red_placeholder_header": true,  // 标题作者占位符标红
    "number_captions": true,         // 图题补 "Fig. N"，表题补 "Table N."
    "normalize_table_refs": true,    // 正文 "Table Ⅰ/I/Ⅱ" → "Table 1/2"
    "reference_heading_unumbered": true
  }
}
```

## 引擎做的事（对应踩过的坑）

| 环节 | 做法 | 为什么 |
|---|---|---|
| 段落原型 | 取模板里**第一个用该样式的段落**，只留 `pPr`，清掉 `w14:paraId` | 复制样式和直接格式，又不复制示例文字；paraId 重复 Word 会报 |
| 正文搬运 | 逐 run 复制并删掉 `rFonts/sz/szCs/kern`，保留 `b/i/vertAlign`，`m:oMath` 整段复制 | 字号字体交给模板样式，斜体上下标和公式不能丢 |
| 标题编号 | 用模板 Heading 样式的 `numPr`，不手写序号 | 模板自带多级编号，手写会重号 |
| 公式 | `tab + oMath + tab + (n)`，居中制表位 + 右对齐制表位 | 与模板示例的"公式居中、编号贴右"一致；别用空格对位 |
| 图 | 按 `cm` 宽度 `add_picture`，图题补 `Fig. N` | 图片关系在新包里重建，不能直接 deepcopy `r:embed` |
| 表 | 按**字体度量**算列宽（PyMuPDF Times-Roman/Bold），`cantSplit` + `tblHeader` | 等比缩放会让 "≤1.0N (%)" 这类表头折行 |
| 参考文献 | 去掉 `[n]` 手写编号，套用模板 `Reference` 样式靠自动编号 | 手写编号 + 自动编号会双份 |
| 收尾 | 按 OOXML schema 顺序重排 `pPr/rPr/trPr/tcPr/tblPr` 子元素 | 顺序错 Word 可能提示"内容有问题" |
| 复核 | Word COM 量页数/表格跨页/图题同页 + PDF 扫 `■` | OOXML 检查看不到分页，必须用真引擎 |

## 两个必须知道的 Word 行为

1. **「与下段同页」(keepNext) 会显示黑色小方块**：文档里只要是 `keepNext`（无论直接格式
   还是样式），Word 打开「显示编辑标记」就在段首画 ■。它**只显示、不打印、不进 PDF**。
   要表格/图题不分页就留着；要编辑视图干净用 `--no-keep-next`（会退回"表格可跨页、
   但断页时表头重复"）。
2. **模板标题样式通常自带 `keepNext`**，所以标题旁的 ■ 是模板固有的，不是本次改动引入。

## 缺内容怎么处理

- 原文标题/作者仍是占位符（含 `use style`、`Given Name`、`XXXXXXXX` 等）：保留模板占位块
  **并以红色标出**，在 `notes` 里写明，不要编造。
- 某章节资料里没有：记录到 `notes`，继续后面的内容。
- 资助声明、作者简介表等模板没有的章节：写进 `drop_styles` / `drop_sections`，并在交付时说明。

## 交付前必做

```powershell
python -m dsh_word retypeset ... --verify --pdf out.pdf      # 版面真测
python -m dsh_word check --docx out.docx --template <模板>    # 确定性规则
```

- `check` 的 `template-conformance`（段落数/样式序列与模板不同）是**预期项**，不必修；
- 交付时给出：`out.docx`、`out.pdf`（Word 导出，别用 LibreOffice 的分页当准）、
  `notes`（跳过了什么、缺了什么）、`issues`（若有）。**不要改用户的原文件**。

## 现成案例

IEEE 双栏会议稿 → IIRM 2026 / TTP 单栏：`--profile ieee-to-ttp`，实测 8 页、
两张表均不跨页、5 张图与图题均同页、PDF 无编辑标记。
