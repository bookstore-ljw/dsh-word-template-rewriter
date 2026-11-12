# Word 流水线面板（client 插件）

`@local/dsh-word-pipeline` 是装在 desktop profile 里的本地 bundle，给右栏加一个
「Word 流水线」面板：资料/模板上传、填充、检查、合批、渲染，以及结果区。
它不重新实现业务逻辑，所有按钮都调用仓库里的 `python -m dsh_word ...`，
和 agent 在终端里跑的是同一套代码。

## 入口（刷新页面后）

三个入口任选：

1. **输入框下方的「▤ Word 流水线」按钮** —— 点一下打开右栏面板；如果当前 DSH 版本的
   右栏标签注册不可用，会自动退化成右侧浮层（同样是一个面板）。
2. **右栏 guide 卡片** —— 打开右栏（右侧边栏开关），guide 里会出现「Word 流水线」卡片，
   点卡片把面板作为标签打开。
3. 直接让 agent 跑（聊天里说一声），面板只是把同一件事按钮化。

## 实时过程（改了什么，能看到）

「填充」不是只返回一个结果，而是把整个过程摊开：

1. `fill` 以 `--progress` 逐步输出 `PROGRESS {json}`，每行对应一次原子操作；
2. Host 边读边记，面板每 350ms 拉一次 `/word-pipeline/progress`，操作列表会一条条冒出来：
   `01 table · table.refs`、`02 text · text.title`…
3. 填充结束后，Host 用**同一个 Word 实例**批量把每一步渲染成 PNG
   （`render-steps`，6 步约 12 秒），存到 `fixtures/build/panel-previews/`；
4. 面板随后**自动回放**：每 0.9 秒切一步，右边预览跟着变；也可以点任意一步，或点「回放」重看。

预览用的是 Word 真实排版（不是内置的 PDF 预览），所以图片尺寸、公式、表格边框都是最终效果。
`/word-pipeline/preview?path=...` 按需渲染并缓存，同一份文档只渲染一次。

## 面板怎么用

| 区域 | 行为 |
|---|---|
| 资料 DOCX | 拖入或点击选择 `.docx`；上传后落到 `uploads/source/`。不传就用 `fixtures/build/source.docx` |
| 模板 DOCX | 拖入或点击选择 `.docx`；上传后落到 `uploads/template/`。不传就用 `fixtures/build/template.docx` |
| 填充 | 有上传资料时走**自动槽位映射**（`--auto-plan`）；否则用 `fixtures/build/plan.yaml` |
| 检查 | 对输出跑全部规则，带模板比对；结果里给 error/warning/info 计数 |
| 合批 | 用 `fixtures/corrections.example.yaml` 跑 Router 合批（示例队列） |
| 渲染 | Word COM → PDF/PNG，产物为 `panel-preview.pdf/png` |

结果区显示：执行的命令、退出码、stdout/stderr、产物绝对路径。所有产物都在
`fixtures/build/` 下。

## 自动槽位映射（`--auto-plan`）

按文档顺序做保守映射，不生成任何文字：

- `text.*` ← 资料里带文字的正文段落，按顺序；
- `formula.*` ← 资料里带 OMML 的段落；
- `image.*` ← 资料里嵌入的图片（写盘到 `assets/source-image-N.ext`）；
- `table.*` ← 资料里的表格。

资料不够时**不静默留空**：CLI 报「未覆盖的槽位」并退出码 2，面板上显示为失败。

## 安全与边界

- Host 端只注册 `/word-pipeline` 前缀路由；随机 token 在页面加载时注入为全局
  `__DSH_WORD_PIPELINE__`，所有请求必须带 `X-Word-Token`，否则 403。
- 只允许操作 `DSH_WORD_ROOT`（默认本仓库）内的路径；上传按类型校验扩展名。
- 命令是白名单（fill/check/apply/render），不会执行任意命令。
- 单次操作 180s 超时；Python 依次尝试 `DSH_WORD_PYTHON` → `python` → `py` →
  本机 `pythoncore-3.12-64\python.exe`。

## 已知限制

- 面板的按钮是「一段操作」，不是对话；需要改内容时还是回聊天或终端。
- 每步预览都要经过 Word：首个文档要等 Word 启动（约 6 秒），之后每步约 1 秒；
  步骤很多的文档，图片会按顺序陆续出现。
- 自动映射只按顺序，不会理解语义；复杂模板请用 `plan.yaml`（面板暂未提供 plan 上传入口）。
- 合批按钮固定用示例修正队列，还没有「在面板里编辑修正」的 UI。
- 页面刷新后才会加载新客户端代码；改 `plugin/` 后需要重新装一次 bundle 再刷新。

## 源码

`plugin/`：`index.js`（Host：路由、上传、调用 CLI）、`client.js`（右栏面板与入口）、
`cordis.patch.yml`、`locale/*`、`icon.svg`。
