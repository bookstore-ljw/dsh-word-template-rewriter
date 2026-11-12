# Agent 管理器（Client 插件）

`@local/dsh-word-pipeline` 现在同时提供三样东西：

1. **Agent 管理器**：左下角的「应用」入口 → 应用抽屉 → Agent 画廊 → 点卡片进入应用；
2. **Word 流水线助手**：资料/模板上传、对话式指令、Word 实时预览；
3. 右栏的轻量「Word 流水线」面板（之前那个，仍然可用）。

## 入口与操作

| 步骤 | 操作 |
|---|---|
| 打开 | 刷新页面（Ctrl+Shift+R）后，点左侧列表里的 **Agent 管理器**（在「插件」下面，同款行格式） |
| 选应用 | 页面里点「Word 流水线助手」这一行（和插件列表同样的行样式） |
| 上传 | 左栏「资料 DOCX」「模板 DOCX」拖入或点击选择；不传就用示例 |
| 指令 | 底部输入框或快捷标签：`填充`、`检查`、`合批`、`渲染`、`把图改成 10cm` |
| 预览 | 右栏实时显示 Word 渲染结果；下方步骤列表可点任意一步回看 |
| 返回 | 页面右上角「← 全部 Agent」 |

> 说明：入口同时注册了左侧导航行（`sidebar.panellist`，id `agents`，order 2）和侧栏底部按钮
> （`sidebar.footer.action`，作为回退）。如果导航行出现，底部那个多余的可以删掉。

图标：自绘 SVG（文档 + 时钟），用于导航行、页面标题和 Agent 卡片；不依赖
`dsh-icon-theme`（该包 0.3.0 要求 dsh `0.2.1-rc.5 || 0.2.2-alpha.3`、0.2.3 要求核心包
`<0.2.0`，与当前 `0.2.2-rc.1` 不兼容，安装被运行时拒绝，未申请版本豁免）。

对话区就是和 Router 的接口：每条指令发到 `POST /word-pipeline/agent/instruct`，
Host 把它翻译成 `python -m dsh_word ...`，把结果、退出码和产物路径作为回复返回。

## Agent 日志与可观测性

每次转换 / 每次对话都创建一条 **run**：

- `runId = <base36 时间>-<随机>`；
- 事件写入 `logs/runs/<runId>.jsonl`，字段
  `{ t, level, phase, message, op?, slot?, ms?, artifact? }`；
- 内存保留最近 60 条运行，`GET /runs?limit=` 列表、`GET /run?id=` 详情；
- 转换的 NDJSON 流里同步下发 `{"type":"log", entry}`，UI 实时追加；
- 预览面板顶部显示当前阶段（`等待运行 / 填充 3/6 / 检查 / 渲染 / 完成`），
  下方「Agent 日志」逐条显示：时间、阶段、消息、耗时（`fill 12083ms`、`check 346ms`、`render 6256ms`）。

![Agent 日志](docs/shots/agent-log.png)

## 文档预览（流式）

转换不再是"等全部完成再出图"，而是 **NDJSON 流式响应**：

- 前端请求 `POST /convert-stream`，服务端边执行边写事件行：
  `{"type":"stage"}`、`{"type":"step"}`、`{"type":"preview"}`、`{"type":"done"}`；
- 前端用 `fetch` + `ReadableStream` 逐行解析，**文本状态和预览图同时流出来**：
  「读取资料… → 填充 2：text.title → … → 检查中… → 渲染中…」；
- 每个原子操作完成即渲染该步 PNG，`preview` 事件一到达就替换右侧预览（带版本号破缓存）；
- 对话改稿（apply）结束后也会重新渲染并更新预览。

实测一次转换过程中：拿到 7 张不同预览帧，状态文本按步骤逐条出现。

预览尺寸：最终渲染 `--dpi 160`（约 1323px 宽），面板里按容器宽度显示（约 500px），
可上下滚动查看整页；之前 `max-height:300px` 会把整页压成缩略图，已移除。

**放大查看**：点击预览图（或双击）打开全屏灯箱，`Esc` / 点背景 / 右上角 `×` 关闭。

![全屏查看](docs/shots/preview-zoom.png)

**修改位置用颜色标出**：每一步的预览帧会先把该操作改动的段落（或表格首行）打成
琥珀色底（`FFF2B2`），再渲染成图，所以流式预览里能直接看到"正在改哪一处"。

![步骤高亮](docs/shots/highlight-step.png)

![预览按面板宽度显示](docs/shots/live-preview.png)

## Word 实时视图（替代 PNG 预览）

- 点「开始转换」会先把模板复制成 `live-preview.docx` 并在 **Word 里打开**；
- 每完成一个原子操作，就把该步骤文档（带黄色高亮）同步进这个 Word 窗口：
  `doc.Content.Delete(); doc.Content.InsertFile(step)`，然后保存；
- 所以 Word 窗口里能看到文档逐步被填、刚改的段落是黄色；
- 转换结束再同步一次干净的最终稿，黄色标记消失；
- 面板右侧不再渲染 PNG，只显示 Word 状态、日志和下载按钮。

实现：`dsh_word/live.py`（`live-open` / `live-sync`）+ Host `/open-word`。

![Word 实时视图](docs/shots/word-live.png)

## 结构映射详情（调试）

「查看映射详情」按钮（Word 卡片下方）会打开一个弹窗，展示本次映射：

- 顶部：模型给出的一句话说明；
- 列表：`模板块 ← 资料块` 与动作（replace / append）；
- 校验问题：类型不符、重复、漏映射、文本过长；
- 「模型原始输出（initial / fix）」可展开看模型的原始 JSON。

数据同时落盘：`logs/mapping/<runId>.json`（含模板/资料结构、prompt、raw、最终映射、校验结果），
接口 `GET /word-pipeline/mapping?id=<runId>`。

![结构映射详情](docs/shots/mapping-modal.png)

## 转换与调整都交给 Agent（后端 Agent = DSH Agent）

「开始转换」和「需要调整？」现在不再调用插件内的小路由，而是：

1. 生成任务说明（模板/资料路径、上次检查问题、要求）；
2. `sessions.create()` 新建一个 DSH 会话，`openSession()` 切过去；
3. 用 `session.prompt()` 直接把说明发给 **DSH 里的 Agent**；
4. 「需要调整」的输入复用同一个会话（sessionId 存 localStorage），继续对话即可。

所以真正干活的是有文件读写、能跑命令、能看 Word 的 Agent，插件只负责生成任务说明和上下文。

![Agent 接手执行](docs/shots/agent-run.png)

如果会话服务不可用，会自动回退到原来的插件内流水线，不会卡死。

## 交给 Agent 处理（真正的 Agent 接手）

插件自带的映射器只是规则 + 单次 LLM 调用，遇到复杂模板仍然会犯傻。面板提供
**「交给 Agent 处理」**：会打开一个新的 DSH 对话，把当前任务写成完整说明预填进去：

- 模板路径、资料路径；
- 上次检查的问题清单；
- 要求：读模板理解章节、按章节从资料取内容、原地替换保留格式、缺内容记录并继续、完成后跑 check 并修复格式。

发送后就是由 **DSH 里的 Agent**（有读写文件、跑命令、看 Word 的能力）直接处理，
而不是插件内的简化调度器。

![交给 Agent](docs/shots/handoff.png)

## 格式自愈

整体检查发现输出与模板的**格式差异**（段落样式、字符格式、页面设置）时，自动执行修复：

1. 把模板对应段落的 `pPr` 复制回输出段落（样式、缩进、行距、编号）；
2. 按段落名恢复 `style`（如 `Heading 2`）；
3. 用模板首个 run 的字体/字号修正输出 run；
4. 恢复页面设置（四边距、纸张大小、方向）；
5. 修复后**重新跑一遍检查**，日志记录「格式自愈：已恢复模板样式」。

实测：`p6: Heading 2→Normal` 被修复，复检后只剩「段落数量与模板不同（10→12）」这类 info，
因为多出来的段落是资料内容，不属于格式问题。

## 与原始模板的对比检查

整体检查时，除了规则校验，还会把**输出文档与原模板逐项对比**：

- 段落数量：`template=10 output=12`；
- 段落样式：`p6: Heading 2→Normal`；
- 字符格式（首个 run 的字体/字号/加粗）；
- 表格数量；
- 页面设置（四边距、纸张大小）。

结果以 `template-conformance` 规则返回，severity 为 `info`，**不阻塞转换**；
即使整体检查通过，面板也会列出这些差异，方便判断"哪里被改动了"。

## 按模板章节逐段填充

模板里的语义标题就是"表单字段"：「标题 / 作者 / 摘要 / 关键词 / 参考文献 / 正文」。
转换时：

1. 按标题把模板切成**有序章节**；
2. 每章从资料里找对应内容（标题取首段、作者取次段、摘要/关键词按标签识别、参考文献按标题后的块）；
3. 逐章写入，每写一块立刻做块级检查；
4. **资料里找不到某章内容时不中断**：记录缺失、在对话里提示「资料里没有找到「关键词」可用的内容，先跳过这部分」，
   继续处理后面的章节；
5. 全部结束后再整体检查；检查不通过也不终止，继续渲染并列出具体规则。

![按章节填充 + 缺失提示](docs/shots/section-fill.png)

## 检查失败不再中断

检查不通过**不会终止转换**：

- 仍然继续渲染，Word 预览与下载都可用（文档带问题，但能打开看）；
- 面板列出失败规则与原因，例如
  `[replacement-character] 文本含 Unicode 替换字符 U+FFFD（乱码）`；
- 日志记录「检查未通过：N 个问题，继续渲染以便查看」；
- 阶段状态：填充 ✅ / 检查 ❌ / 渲染 ✅。

![检查失败仍可查看](docs/shots/check-failed.png)

## 结构映射（章节 + 顺序 + 审校）

- **章节树**：`analyze` 输出 `sections`（heading → 子块 id），模型按章节语义对齐，
  而不是简单按段落顺序；
- **全文**：每个块带完整文本（最多 4000 字）参与映射，不再只看 120 字预览；
- **保持顺序**：图片/公式/表格用 `insert-after` 锚定到"它前面最近的文本块所映射的模板块"，
  连续插入按源顺序排队，不再统一堆到文末；
- **语义审校**：第二次 LLM 调用逐条判断"这段内容放这里是否合理"，
  有问题时把问题清单回喂模型重新映射一次；
- **规则校验**：资料内容缺映射为 error，类型/重复/超长为 warn。

映射结果与两次 LLM 的原始输出都能在「查看映射详情」和 `logs/mapping/<runId>.json` 里看到。

> 注意：默认模型是推理模型，映射 + 审校一次完整转换约 2 分钟（单次 LLM 约 100 秒）。
> 需要更快可以下调 `maxTokens` 或对小文档跳过审校。

## 模板没有占位符时：原地替换

模板保留自己的结构，内容**原地替换**，不做整段清空重建：

1. 标题段落（Heading/标题样式）原样保留；
2. 其余正文段落按顺序接收资料内容，**保留每个段落的 `pPr` 与首个 run 的 `rPr`**，
   所以段落样式、缩进、行距、字体都不变；
3. 资料比段落多时，复制最后一个正文段落的格式继续追加；
4. 资料比段落少时，多余的示例段落**只清空文字、不删除段落**，格式仍然在；
5. 图片、公式、表格同样复制进对应段落。

![Word 实时视图](docs/shots/word-live.png)

## Word 保存策略

实时同步每完成一步都会更新 Word 窗口内容，但**不会每步都保存文件**：

- 文本块：只改当前段落（`Paragraph.Range.Text`），**不清空全文**，所以 Word 不会跳回第一页；
- 每步都会：清掉旧高亮 → 选中刚改的段落 → `ScrollIntoView` 滚到它 → 黄色高亮；
- 图片/公式/表格没有段落级定位时，才回退整篇同步，并同样跳到对应段落；
- 距上次保存不足 30 秒的更新只改屏幕内容（`--no-save`）；每满 30 秒保存一次；
- 不再在结束时整篇重写，避免视图重置。

## 对话（Agent 调度）

「需要调整？」输入框现在**真的走模型**，不再只是关键词匹配：

1. 前端 `POST /agent/chat`（带最近几条用户消息 + 当前资料/模板路径）；
2. Host 用 `llm.stream()` 调 **DSH 当前默认模型**
   （`agentDefaultModel.currentSelection()`）做路由，系统提示要求模型只输出一个 JSON：
   `{"action":"convert|fill|check|render|set_image_width|none","width":…,"reply":"…"}`；
3. Host 按 action 执行对应 CLI（与按钮同一套 `runConvert` / `executeOp`），把结果和回复返回前端；
4. 前端把 `convert` 的结果直接渲染成阶段卡 + 下载按钮 + 预览。

模型不可用或调用失败时返回 `fallback: true`，前端自动退回本地指令解析，并在回复后标注
「模型未接入，已用本地指令解析」。

例子：输入「帮我把资料转成成稿」→ 模型判定 `convert` → 填充 6 步 / 检查 / 渲染全绿，
右侧出预览，下方出现「下载 DOCX / 下载 PDF」。

![对话驱动转换](docs/shots/agent-chat.png)

## 一键转换（主流程）

> **模板必须带占位符**：转换只做"把资料复制进 `{{ ... }}` 槽位"，不会改写模板文字。
> 模板里一个 `{{ }}` 都没有时，管线不会做任何填充（旧版本会静默成功，现在会明确报错）。

打开 Agent → 「Word 流水线助手」：

1. 拖入或点击上传资料 / 模板（留空即用示例）；
2. 上传框本身就是状态：上传后框内标题变为 **「已加载文档」**、下面显示文件名，
   框下方有一条 **上传进度条**（真实 XHR 进度，完成后 100%），右侧 `×` 可清除改用示例；
3. 选择结果保存在 `localStorage`（键 `dsh-word-files-v1`），刷新页面/重启应用后仍在；
4. 点 **开始转换**——一次点击依次执行 `fill → check → render`；
3. 完成后显示三步状态（绿点）与 **下载 DOCX / 下载 PDF** 按钮；
4. 右侧自动显示 Word 渲染的成品和每一步的中间版本。

只有「哪里不合适」时才用下面的对话区（例如「把图改成 10cm」），它会跑修正并重新生成。
Host 端新增：`POST /word-pipeline/convert`（串联三个阶段）、
`GET /word-pipeline/download?path=…&t=…`（带 `Content-Disposition` 的文件下载）。

## 支持的指令

| 输入 | 行为 |
|---|---|
| `填充` / `fill` | 用当前模板 +（已上传的）资料跑自动映射填充，返回 6 个原子操作 |
| `检查` / `check` | 跑全部确定性规则 + 模板格式比对，回复里带 error/warning/info |
| `合批` / `apply` | 用示例修正队列跑 Router 合批 |
| `渲染` / `render` | Word COM → PDF/PNG |
| 含 `10cm` / `10厘米` 的句子 | 把上一次填充里的 image 槽位改成该宽度，重新合批并刷新预览 |

不是万能 NLP：它按关键词和宽度正则路由，超出范围的输入会回复可用指令。

## 架构

```
Client（本插件）
  ├─ 入口：sidebar.footer.action#agent-manager
  ├─ 应用：shell.overlay#agent-manager-app（全屏应用窗口）
  └─ 右栏：sidebar.right.pane.tab#word-pipeline
        │ fetch（X-Word-Token）
        ▼
Host（本插件 index.js）
  ├─ POST /word-pipeline/agent/instruct   指令路由
  ├─ POST /word-pipeline/op               直接操作（fill/check/apply/render）
  ├─ POST /word-pipeline/upload           文件落盘到 uploads/
  ├─ GET  /word-pipeline/progress         原子操作进度
  └─ GET  /word-pipeline/preview          Word 渲染 PNG（缓存）
        │ spawn
        ▼
python -m dsh_word （仓库里的确定性 Worker）
```

每个 Agent = 一张卡片 + 自己的一组端点 + 自己的界面。再加一个 Agent 时，照这个结构
复制一份端点与卡片即可，彼此不共享状态。

## 安装 / 重装

插件目录：`plugin/`，已经装进 `desktop` profile。改完代码后：

```powershell
# 重新装（会覆盖链接，重新建立 bundle 行）
# 由 agent 调用 plugin_manager install_bundle，target = D:\desktop\project\DSH-word\plugin
```

然后**刷新页面**。Host 侧改动需要重载插件行（agent 用 `set_plugin` 关再开），
Client 侧改动刷新页面即可。

## 设计规范（DSH primitives tokens）

界面按 DSH 自带组件的观感重做：卡片用 `--dsw-shadow-lv1` + `--dsw-radius-lg`，行高/按钮/
输入框用 `--dsw-radius-sm`，标题字号用 `--dsw-font-l-20-*`，正文 `--dsw-font-s-14-*`，
辅助文字 `--dsw-font-xs-13-*` / `--dsw-font-xxs-12-*` / `--dsw-font-xxxs-11-*`，颜色全部来自
`--dsw-alias-*`（label / bg-layer / border / brand / interactive / state）与 `--dsw-specific-*`
（气泡、输入框、文档预览底）。没有写死的 hex。

一处必须说明：DSH **没有 spacing 类 token**（扫描了 414 个 `--dsw-*` 变量，只有
radius / font / shadow / color / elevation）。所以内边距与外边距仍是 4px 栅格的数值
（4/8/12/16/20/24/32），这是唯一没能 token 化的部分；控件尺寸（图标块 40px、预览图标 56px）
同理。若后续 DSH 增加 spacing tokens，替换这几处即可。

### 字体系统

- **只使用系统原生无衬线字体栈**：
  `-apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB",
  "Microsoft YaHei", "Helvetica Neue", Arial, "Noto Sans", sans-serif`
  （Windows 落在 Segoe UI + 微软雅黑，macOS 落在 SF Pro + 苹方）；
- 不再引用 `--dsw-font-family-brand`（Montserrat 未安装时会回退，字重/字宽不稳定）；
- 字重只用 400 / 600（此前有 650 这种非标准值，浏览器会做合成加粗，导致发虚）；
- 关闭 `-webkit-font-smoothing: antialiased` 与 `text-rendering: optimizeLegibility`，
  恢复系统默认的次像素渲染；
- 字距统一 0（此前 `.01em/.02em/.2px` 对中文会造成轻微错位与模糊）；
- 字阶（全部来自 `--dsw-font-*`）：
  | 用途 | Token |
  |---|---|
  | 页头标题 | `--dsw-font-xl-24-*`（24/32，600） |
  | Agent 卡片标题 | `--dsw-font-m-18-*`（16/28，600） |
  | 面板标题 | `--dsw-font-xs-strong-13-*`（13/20，600） |
  | 正文 / 对话 | `--dsw-font-s-14-*`（14/22） |
  | 次要说明 | `--dsw-font-xs-13-*`（13/20） |
  | 元信息 | `--dsw-font-xxs-12-*`（12/18） |
  | 步骤 chips | `--dsw-font-xxxs-11-*` + `--ds-font-family-code`（等宽） |

### 视觉基调（高级感）

- 顶部是一条**深色渐变横幅**：`neutral-bluish-1000 → deepseek-800 → violet-600`，
  白色标题与副标题，右上角状态胶囊/返回按钮用半透明白；
- 主按钮「开始转换」与「下载 DOCX」用 `deepseek-500 → violet-600` 渐变 + 紫色投影；
- **两块面板不再是白色**：左「一键转换」= 8% deepseek 蓝 + 白，右「文档预览」= 9% 紫 + 白，
  预览底板再叠 5% 紫；上传框/输入框仍是白色，作为可输入的控件；
- 结果卡片用 5% 紫色底 + 16% 紫边，成功点仍是绿色；
- 页面画布是「紫色/蓝色径向光晕 + 模块底色」；
- 面板顶部加 2px 紫色描边，卡片保留 `elevation-panel`。

### 边框与画布

`--dsw-alias-border-l1` 在浅色皮肤下只有约 4% 黑，卡片边等于没有。现在：

- 卡片描边：`--dsw-alias-settings-card-stroke`（= `border-l4`，约 16% 黑）
- 卡片阴影：`--dsw-elevation-panel`（0.5px 描边 + 两层投影）
- 页面画布：`--dsw-alias-bg-module-platform`（浅灰），与白色卡片形成层次
- 面板标题分隔：`--dsw-alias-border-l2`；上传区虚线：`border-l4`

这样卡片边界在浅色/深色下都能看出来，不再和背景糊在一起。

### 图标与强调色

两个图标分开，避免「管理器」和「Agent」看起来是同一个东西：

- **Agent 管理器**（应用外壳）：紫色圆角底 `--dsw-static-violet-600` + 白色 2×2 网格，
  用于左侧导航行和列表页页头；
- **Word 流水线助手**（Agent）：蓝色圆角底 `--dsw-static-deepseek-500` + 白色文档 +
  琥珀色管线节点 `--dsw-static-amber-400`，用于 Agent 卡片、工作区页头、预览空态；
- 强调色统一走 `--dsw-static-deepseek-*`：主按钮 / 下载 DOCX / 上传图标 / 用户气泡 /
  选中步骤；成功状态点走 `--dsw-alias-state-success-primary`；原来 `brand-primary` 在
  浅色皮肤下接近黑色，缺少区分度，所以显式改用 deepseek 蓝；
- 中性面仍用 `--dsw-alias-bg-layer-*` / `border-*`，保证与宿主一致。

## 入口清理

- 删除了输入框右下角的「▤ Word 流水线」启动器（`conversation.composer.dock` 注册）；
- 删除了侧栏底部的重复入口，只保留「插件」下面的导航行；
- 卸载了 `dsh-dream-skin`，恢复官方默认皮肤。

## 可视化验证（@feather_wch/dsh-plugin-ui-debug）

桌面版走 Electron `dsh-app://` + IPC，HTTP 端口带 token 鉴权，没法直接截。验证用了一个
临时 web profile：

```powershell
dsh --profile uiweb --from-default-profile web --dump-config   # 建一个干净 web profile
dsh plugin --profile uiweb add D:\desktop\project\DSH-word\plugin
dsh --profile uiweb --no-open                                  # 启动后会打印 http://127.0.0.1:PORT/?token=...
```

拿到带 token 的 URL 后，用 `ui_shot` 截图、`ui_drive` 点「Agent 管理器」→ 卡片 → 输入指令，
全程真实 Chrome + Playwright。实测确认并修掉了三个问题：

1. 工作区没铺满：外层多包了一层 `.wp-workspace`，改为直接渲染 → 两栏 505/569px、高度铺满；
2. 步骤 chips 文字被截断：继承了旧 `.wp-step` 的 `overflow:hidden`，改为 `overflow:visible`；
3. 侧栏底部多了一个重复入口：确认导航行注册成功后删除了 footer 回退按钮。

验证完停止该 profile 的服务；profile 目录保留，后续要再看界面时重复上面三步即可。

## 边界

- 只内置了一个 Agent（Word 流水线助手）；「更多 Agent」是结构上的预留，不是已实现的功能。
- 没有用 iframe：界面直接以 React 渲染在 `shell.overlay` 里，这样能继承 DSH 主题与字号。
  如果某个 Agent 是独立的 Web 服务，同一张卡片里换成 `<iframe src=...>` 即可。
- 对话不是 LLM 对话，是命令路由；要真正"聊天改文档"，需要把 Router（模型）接进来。
- 预览依赖 Word COM，首次约 6 秒，之后每步约 1 秒。
