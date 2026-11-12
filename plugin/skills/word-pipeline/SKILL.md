---
name: word-pipeline
description: Word 流水线助手：把资料 DOCX 的内容按模板章节填进模板 DOCX（保留模板格式），逐块检查、格式自愈、整体校验并渲染预览。通过 dsh_word 命令行与 Word COM 工作，适合"资料 + 模板 → 成稿"的确定性替换任务。
whenToUse: 用户要求把资料内容填进 Word 模板、按模板章节生成成稿、检查/修复 Word 版式、或在 Word 里实时查看填充过程时使用。
---

# Word 流水线助手（agent 版）

这套技能把"填模板"拆成确定性步骤，Agent 负责理解结构与决策，Python 负责执行。

## 工作区约定

- 根目录：`D:\desktop\project\DSH-word`（或 `DSH_WORD_ROOT`）；
- 上传：`uploads/source/`、`uploads/template/`；
- 产物：`fixtures/build/out.docx`、`panel-preview.pdf/png`；
- 实时 Word：`fixtures/build/live-preview.docx`；
- 映射调试：`logs/mapping/<runId>.json`；
- 运行日志：`logs/runs/<runId>.jsonl`。

## 第 1 步：拆结构

```powershell
python -m dsh_word analyze --docx <模板.docx>
python -m dsh_word analyze --docx <资料.docx>
```

输出 JSON：`blocks`（heading / paragraph / empty / image / formula / table，含 `id`、`section`、全文）
与 `sections`（heading → 子块）。**先读 sections，再读 blocks**，理解每章要什么。

## 第 2 步：做映射（Agent 决策）

生成 `mapping.json`：

```json
{"mapping":[{"template":"paragraph:1","source":"paragraph:0","action":"replace"}],"notes":"一句话"}
```

- 模板 `heading` 不映射；
- 文本用 `replace`；
- 图片/公式/表格保持资料顺序：用 `{"template":"<前一个文本块 id>","source":"image:3","action":"insert-after"}`；
- 没有合适锚点用 `{"template":"append","source":"table:0","action":"append"}`；
- 每个资料块最多用一次。

## 第 3 步：校验映射

```powershell
python -m dsh_word map-verify --template <模板> --source <资料> --mapping mapping.json
```

`ok:false` 时按 `issues` 修正 mapping；类型不符、重复、漏映射（资料内容未用）都会报。

## 第 4 步：执行填充

```powershell
python -m dsh_word map-apply --template <模板> --source <资料> --mapping mapping.json --out fixtures/build/out.docx --steps-dir fixtures/build/panel-steps --progress
```

`PROGRESS` 每块一条（可用于 Word 实时同步）；`CHECK` 是块级检查（空文本 / 残留 `{{ }}` / 替换字符 / 图片 / 公式 / 表格）。

## 第 5 步：整体检查与自愈

```powershell
python -m dsh_word check --docx fixtures/build/out.docx --template <模板>
python -m dsh_word heal --template <模板> --docx fixtures/build/out.docx
```

- `check` 失败不阻断：记录规则与位置；
- 与模板的样式/字体/页边距差异可用 `heal` 恢复；恢复后再 `check` 一次。

## 第 6 步：渲染与实时 Word

```powershell
python -m dsh_word render --docx fixtures/build/out.docx --pdf fixtures/build/panel-preview.pdf --png fixtures/build/panel-preview.png --dpi 160
python -m dsh_word live-open --file fixtures/build/live-preview.docx
python -m dsh_word live-sync --target fixtures/build/live-preview.docx --source <步骤.docx> [--no-save]
```

实时同步规则：文本块只改当前段落并跳转高亮；每 30 秒保存一次。

## 缺内容怎么办

某章在资料里找不到对应内容时：**记录缺失、在对话里说明、继续后面的章节**，不要停下来或编造内容。

## 输出给用户

最终给出：`out.docx`、`panel-preview.pdf` 路径；检查问题列表（若有）；缺失章节列表（若有）。
