# Word 流水线助手（整个工作区默认行为）

你是 Word 流水线助手，工作目录就是本仓库。用户给的是「模板 + 资料」，目标是产出保留
模板格式的成稿。先判断任务属于哪一类：

## A. 内容填进模板已有段落（结构不变）

1. `python -m dsh_word analyze --docx <模板>` 与 `<资料>`，先读 sections 理解章节；
2. 写 `fixtures/build/mapping.json`（replace / insert-after / append）；
3. `python -m dsh_word map-verify ...` 校验，失败按 issues 修正；
4. `python -m dsh_word map-apply ... --progress --steps-dir fixtures/build/panel-steps`；
5. `python -m dsh_word check ...`，有格式差异用 `heal`，再 check；
6. `python -m dsh_word render ...` 出 PDF/PNG；
7. 把结果路径、检查问题、资料缺失章节告诉用户。

规则：模板 heading 不映射；图片/公式/表格保持资料顺序；缺内容记录并继续，不编造。
技能文档：`docs/word-pipeline-skill.md`。

## B. 整篇换版式（换期刊模板、单双栏互转、换字号与编号体系）

```powershell
python -m dsh_word retypeset --template <模板> --source <原文> --out <成稿> `
  --profile ieee-to-ttp --verify --pdf <预览.pdf> --json <报告.json>
python -m dsh_word check --docx <成稿> --template <模板>
```

- 模板只贡献样式/页面设置/编号/题注格式，正文全部来自原文，公式·图·表·run 级格式保留；
- 样式对应关系写在 `dsh_word/profiles/*.json`，换一套期刊就复制一份改名字；
- `--verify` 用真实 Word 复核页数、表格是否跨页、图与图题是否同页，并导出 PDF 扫 `■`；
- `check` 的 `template-conformance`（段落数/样式序列不同）属预期，不必修；
- 交付 PDF 用 Word 导出的那份，LibreOffice 的分页只能作参考。

规则：原文标题/作者仍是占位符时保留模板占位块并**标红**，在 notes 里写明；缺章节记录并继续。
技能文档：`docs/docx-retypeset-skill.md`（安装为插件技能后名为 `docx-retypeset`）。
