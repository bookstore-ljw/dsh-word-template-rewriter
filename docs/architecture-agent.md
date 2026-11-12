# Agent 版架构（UI 不变，后端是 DSH Agent）

## 1. 变化点

旧版：插件内自带"LLM 映射 + 规则流水线"，能力上限受插件实现限制。

新版：**插件只保留 UI 与任务下发**；理解模板、决定映射、处理缺失、决定自愈，全部由
**DSH 里的 Agent**（有读写文件、跑命令、看结果的工具）完成，Python 引擎只提供确定性命令。

```
┌──────────────────────────────────────────────────────────────┐
│ Web UI（client.js，界面保持不变）                             │
│  上传卡 · 开始转换 · Word 实时视图 · Agent 日志 · 下载        │
└───────────────┬──────────────────────────────────────────────┘
                │ 任务说明（模板/资料路径 + 要求 + 上次问题）
                ▼
┌──────────────────────────────────────────────────────────────┐
│ DSH Agent（真正的模型 + 工具）                                │
│  · skill: word-pipeline（本仓库 plugin/skills/）              │
│  · 工具：读/写文件、bash、office、web…                        │
└───────────────┬──────────────────────────────────────────────┘
                │ python -m dsh_word <cmd>
                ▼
┌──────────────────────────────────────────────────────────────┐
│ Python 引擎（确定性能力，不调用模型）                          │
│  analyze · map-apply · map-verify · section-fill · check      │
│  heal · render · live-open / live-sync                        │
└───────────────┬──────────────────────────────────────────────┘
                ▼
          Word / 产物（out.docx · live-preview.docx · logs）
```

## 2. UI 与 Agent 的对应关系

| UI 元素 | 现在的行为 |
|---|---|
| 「开始转换」 | 生成任务说明，新建 DSH 会话并 `prompt()` 给 Agent |
| 「需要调整」输入 | 复用同一会话，追加要求发给 Agent |
| 「在 Word 中打开」 | 仍由插件调用 `live-open` |
| 「查看映射详情」 | 读取 `logs/mapping/<runId>.json` |
| Agent 日志 | 插件 run ledger（转换时）；Agent 执行时看对话 |

## 3. 工具契约（Agent 看到的命令）

| 命令 | 作用 |
|---|---|
| `analyze --docx` | 结构块 + 章节树 + 全文 |
| `map-verify` | 校验映射（类型/重复/漏映射） |
| `map-apply` | 按映射原地替换，输出步骤与块级检查 |
| `section-fill` | 按模板标题章节逐段填充（缺内容上报） |
| `check` | 整体规则 + 与模板对比 |
| `heal` | 恢复模板样式/字体/页边距 |
| `render` | PDF/PNG |
| `live-open` / `live-sync` | Word 实时同步（跳转 + 黄色高亮，30 秒保存一次） |

## 4. 一个转换的完整链路

```
点「开始转换」
  → 插件生成 brief（路径 + 要求 + 上次问题）
  → sessions.create() + prompt(brief)
  → Agent 加载 skill: word-pipeline
  → analyze 模板/资料 → 读 sections 理解章节
  → 写 mapping.json → map-verify
  → map-apply（逐块 PROGRESS/CHECK，同步 live-preview）
  → check → heal（有格式差异时）→ check
  → render → 回报 out.docx / 问题 / 缺失
```

## 5. 为什么这样更聪明

- 理解与决策交给模型，规则只做执行与校验；
- 缺内容、映射不合理、格式漂移都能在对话里解释与重试；
- 插件不再维护第二套"迷你 Agent"，减少行为分叉。

## 6. 扩展

- 新能力 → 新增 `dsh_word` 子命令，并在 skill 里登记；
- 新 Agent 卡片 → UI 增加入口，任务说明里带上对应 skill 与命令；
- 调试 → `logs/runs/*.jsonl`、`logs/mapping/*.json`。
