# 安装与复用

三种用法，按需要选一种。核心能力就一条命令：

```bash
dsh-word retypeset --template <模板.docx> --source <原文.docx> --out <成稿.docx> \
                   --profile ieee-to-ttp --verify --pdf <预览.pdf> --json <报告.json>
```

> `--verify` 需要 Windows + 已安装 Microsoft Word（调用 COM 量页数、查表格是否跨页、
> 图与图题是否同页，并导出 PDF）。没有 Word 时去掉 `--verify --pdf`，其余流程照常。

---

## 方式一：装成 Python 包（最省事，任何平台）

```bash
# 直接装最新代码
pip install "git+https://github.com/USER/DSH-word.git"

# 想要 Word 版面复核（Windows）
pip install "git+https://github.com/USER/DSH-word.git#egg=dsh-word[render]"
```

安装后即可使用：

```bash
dsh-word retypeset --template 模板.docx --source 原文.docx --out 成稿.docx
dsh-word analyze --docx 原文.docx          # 看原文都用了哪些样式名
dsh-word check --docx 成稿.docx --template 模板.docx
```

## 方式二：克隆后一键安装脚本

Windows（PowerShell）：

```powershell
git clone https://github.com/USER/DSH-word.git
cd DSH-word
pwsh -ExecutionPolicy Bypass -File .\scripts\install.ps1 -WithRender
```

macOS / Linux：

```bash
git clone https://github.com/USER/DSH-word.git
cd DSH-word
bash scripts/install.sh --render
```

脚本做三件事：`pip install`、把 `plugin/skills` 下的技能注册进检测到的 DSH 配置目录、
跑一遍单元测试自检。

## 方式三：作为 DSH 插件（让 Agent 自己会用）

本仓库同时是一个 DSH 插件包（`plugin/`），`plugin/skills/` 下带两个技能：

| 技能 | 用途 |
|---|---|
| `docx-retypeset` | **整篇换版式**（换期刊模板、单双栏互转、换字号与编号体系） |
| `word-pipeline` | 内容填进模板已有段落槽位（结构不变） |

让 DSH 识别插件的方式（任选其一）：

```powershell
# A. 直接把插件目录链进配置（本机开发时最常用）
New-Item -ItemType Junction -Path "$env:USERPROFILE\.dsh\profiles\<profile>\node_modules\@local\dsh-word-pipeline" -Target "<仓库>\plugin"

# B. 用插件面板安装本地包
#    插件管理 → 安装 bundle → 选择 <仓库>\plugin
```

装好后重开会话，技能列表里会出现 `docx-retypeset`。

---

## 五分钟上手：把一篇 IEEE 双栏稿改成单栏期刊格式

```powershell
# 1. 看两份文件各自的样式名
dsh-word analyze --docx 原文.docx
dsh-word analyze --docx 模板.docx

# 2. 跑重排（内置 ieee-to-ttp 对应表）
dsh-word retypeset --template 模板.docx --source 原文.docx --out 成稿.docx `
  --profile ieee-to-ttp --verify --pdf 预览.pdf --json 报告.json

# 3. 看报告：notes 里有"跳过了什么/缺了什么"，layout.issues 为空即版面 OK
Get-Content 报告.json
```

`报告.json` 关键字段：

```jsonc
{
  "heading": 10, "paragraph": 35, "equation": 5, "figure": 5, "table": 2, "reference": 14,
  "notes": ["资料标题是占位符（…），保留模板占位标题并标红", "题注补编号：Fig. 1", "跳过章节：Authors' background"],
  "styles_missing": [],
  "layout": { "page_count": 8, "tables": [{"index":1,"rows":9,"split":false}],
              "figures": [{"index":1,"same_page":true}], "issues": [], "pdf_editor_marks": 0 }
}
```

## 换一套期刊模板

样式名不同就复制一份 profile 改名字，**不用改代码**：

```bash
cp dsh_word/profiles/ieee-to-ttp.json dsh_word/profiles/ieee-to-elsevier.json
# 编辑 target / source 两个段落：左边填模板的样式名，右边填原文的样式名
dsh-word retypeset ... --profile ieee-to-elsevier
```

字段含义见 [docs/docx-retypeset-skill.md](docs/docx-retypeset-skill.md)。

## 常见问题

| 现象 | 原因 / 处理 |
|---|---|
| 段首出现黑色小方块 ■ | Word「显示编辑标记」对带「与下段同页」的段落画的标记，**不打印、不进 PDF**。不想看到就加 `--no-keep-next`（代价：表格可能跨页，但断页时表头会重复） |
| 表格跨页 | 默认已开「与下段同页 + 行不跨页 + 表头重复」；若仍跨页说明该页确实放不下，可调小 `--figure-width-cm` 或加 `--table-width-twip` |
| 表头折行 | 装了 `pymupdf` 会按真实字宽分配列宽；没装则退化为等比缩放 |
| `check` 报 `template-conformance` | 跨模板重建时段落数与模板必然不同，属**预期**，不是错误 |
| PDF 页数与预期不符 | 以 Word 导出的 PDF 为准，LibreOffice 分页只能参考 |
| 中文乱码 | 全库 UTF-8；Windows PowerShell 5.1 请用 `Get-Content -Encoding utf8` |

## 依赖

| 依赖 | 必需 | 用途 |
|---|---|---|
| python-docx / lxml | 是 | OOXML 读写 |
| Pillow | 是 | 图片尺寸 |
| PyYAML | 是 | 检查规则 |
| pywin32 | 可选（Windows） | Word COM：版面复核、PDF 导出 |
| pymupdf | 可选 | 表格列宽按字宽计算、PDF 页面 raster |

## 许可证

MIT，见 [LICENSE](LICENSE)。
