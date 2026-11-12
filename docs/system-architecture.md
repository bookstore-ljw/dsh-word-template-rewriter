# 完整 Agent 系统架构（四部分整合）

> 目标：一个前端 UI（我们的「Word 流水线助手」面板）＋ 一个 Agent 后端（DSH Agent，
> 可选 Word Master 预设）＋ 一套确定性 Word 引擎（dsh_word）＋ 一个全局观测台（dsh-office），
> 四者组成一个完整的 Agent 应用，而不是四个互不相干的插件。

## 1. 分层

```
┌────────────────────────────────────────────────────────────────────┐
│ ① 前端 UI                                                          │
│  Word 流水线助手（client.js，Agent 管理器面板）                     │
│   上传 · 开始转换 · 调整输入 · Word 实时视图 · Agent 日志 · 下载     │
│  dsh-office（overlay）                                              │
│   办公室总览：workspace / session / token / subagent / 日志         │
└───────────────┬───────────────────────────────┬────────────────────┘
                │ 面板请求（/word-pipeline/*）   │ 会话事件 / 全局观测
                ▼                               ▼
┌────────────────────────────────────────────────────────────────────┐
│ ② DSH 平台（源码层）                                                │
│  sessions / agent-loop / llm / tools / webserver / authorization    │
│  · 会话生命周期与事件流（assistant / tool / turn）                  │
│  · 模型调用与工具审批                                                │
└───────────────┬───────────────────────────────┬────────────────────┘
                │ Agent 执行                     │ 插件注册
                ▼                               ▼
┌──────────────────────────────────────┐  ┌──────────────────────────┐
│ ③ Agent 层                            │  │ ④ 后端插件（Host）        │
│ 默认 Agent 或 Word Master 预设         │  │ 路由 /word-pipeline/*     │
│  · word_* 工具（dsh-word-master）      │  │ run ledger · 引擎子进程   │
│  · skill: word-pipeline（本仓库）      │  │ Word 同步 · 预览          │
│  · 通用工具：fs / pwsh / web / todo    │  │ Python 引擎 dsh_word/     │
└──────────────────────────────────────┘  └──────────────────────────┘
```

## 2. 四条工作流

### A. 一键转换（面板内确定性）
`开始转换 → /convert-stream → runConvert（fill/section-fill → check → heal → render）
→ 流式 stage/step/preview/log → 面板展示 → 下载`

### B. 交给 Agent（智能兜底）
`交给 Agent 处理 / 开始转换(委托模式) → sessions.create({agentPreset:'word-master'})
→ prompt(任务说明) → Agent 调 word_* 或 dsh_word 命令 → 结果写入 fixtures/build/out.docx
→ 面板刷新预览与检查项`

### C. 观测（dsh-office）
`office 面板读取 DSH 会话与 token → 展示每个 Agent 在做什么；
我们的 run ledger（logs/runs）作为补充，展示转换阶段与耗时。`

### D. 技能（word-master + word-pipeline）
`选中 Word Master 预设 → 自动加载 word-master skill；
本仓库的 docs/word-pipeline-skill.md 由任务说明引用，确保 Agent 知道 dsh_word 命令。`

## 3. 集成契约

| 方向 | 接口 | 说明 |
|---|---|---|
| UI → Host | `/word-pipeline/convert-stream` | NDJSON：stage / step / preview / log / done |
| UI → Host | `/word-pipeline/agent/run-stream` | 面板内 Agent 工具循环（实验） |
| UI → Agent | `sessions.create` + `session.prompt` | 委托真 Agent（可带 `agentPreset`） |
| Host → 引擎 | `python -m dsh_word <cmd>` | analyze / map-apply / check / heal / render / live |
| Host → Office | 无（office 直接读 DSH 会话） | 同一数据源，零耦合 |
| Agent → 引擎 | 同上（bash 工具） | 技能文档给出命令契约 |

## 4. 数据与产物

```
uploads/{source,template}/  用户上传
fixtures/build/out.docx     成稿
fixtures/build/live-preview.docx  Word 实时同步副本
fixtures/build/panel-preview.pdf/png
logs/runs/<runId>.jsonl     转换账本
logs/mapping/<runId>.json   映射调试
```

## 5. 各部分的改造点

| 部分 | 现状 | 整合后 |
|---|---|---|
| Word 流水线助手 | 面板 + 自建流水线 + 委托按钮 | 面板为唯一入口；委托时带 `agentPreset: word-master` |
| dsh-word-master | Agent 预设 + `word_*` 工具 | 作为"写作/规范/质检"后端，被委托会话使用 |
| dsh-office | 全局办公室面板 | 作为系统观测台，查看 Agent 会话/token/日志 |
| DSH 源码 | 平台 | 提供 sessions/llm/tools/审批；不修改源码 |

## 6. 当前选择：B 模式（Word Master 会话）

- 「开始转换」→ `sessions.create({ agentPreset: 'word-master' })` → `prompt(任务说明)`；
- 会话默认权限预设改为 `word-master-auto`（`sandbox: danger-full-access`、`approval: never`），
  避免 Agent 调 shell 时卡在「待审批」；
- 该默认预设对**所有新会话**生效；如需收紧，删掉 `defaultPreset` 或改回 `workspace-write`；
- 会话服务不可用时回退到面板内确定性流水线。

## 7. 待决策

1. **委托会话是否在面板内流式展示**（不跳转）——需要解决工具审批（headless 会话会卡在"待审批"）；
2. **面板的转换按钮走哪条后端**：确定性流水线（稳）还是 Word Master 预设（强但慢、需审批）；
3. **dsh-office 入口**：保留右下角 FAB，或在面板里加"打开办公室"按钮。
