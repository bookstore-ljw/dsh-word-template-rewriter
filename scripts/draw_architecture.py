"""Draw the current Word pipeline agent architecture as a PNG."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

W, H = 1680, 1120
BG = "#FFFFFF"
INK = "#0F1115"
MUTED = "#5B6470"

FONT_DIR = Path("C:/Windows/Fonts")
FONT = FONT_DIR / "msyh.ttc"


def font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONT), size)


TITLE = font(34)
H2 = font(24)
BODY = font(17)
SMALL = font(15)


def box(draw, x, y, w, h, title, lines, fill, border):
    draw.rounded_rectangle([x, y, x + w, y + h], radius=14, fill=fill, outline=border, width=2)
    draw.text((x + 20, y + 14), title, font=H2, fill=INK)
    ty = y + 52
    for line in lines:
        draw.text((x + 20, ty), line, font=BODY, fill=MUTED)
        ty += 26


def arrow(draw, x1, y1, x2, y2, label=None, color="#4B5563"):
    draw.line([x1, y1, x2, y2], fill=color, width=3)
    if abs(y2 - y1) < 6:  # horizontal
        if x2 > x1:
            draw.polygon([(x2, y2), (x2 - 13, y2 - 8), (x2 - 13, y2 + 8)], fill=color)
        else:
            draw.polygon([(x2, y2), (x2 + 13, y2 - 8), (x2 + 13, y2 + 8)], fill=color)
        if label:
            draw.text(((x1 + x2) / 2 - 50, y1 - 26), label, font=SMALL, fill=color)
        return
    if y2 > y1:  # downward
        draw.polygon([(x2, y2), (x2 - 8, y2 - 13), (x2 + 8, y2 - 13)], fill=color)
    else:  # upward
        draw.polygon([(x2, y2), (x2 - 8, y2 + 13), (x2 + 8, y2 + 13)], fill=color)
    if label:
        draw.text(((x1 + x2) / 2 + 12, (y1 + y2) / 2 - 12), label, font=SMALL, fill=color)


img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

d.text((70, 34), "Word 流水线助手 · Agent 架构（当前实现）", font=TITLE, fill=INK)
d.text((70, 78), "DSH 客户端插件 → Host 编排 → Python 确定性引擎 → Word / 产物", font=BODY, fill=MUTED)

# ---------------- layers ----------------
box(
    d, 70, 130, 1080, 150,
    "1. Web UI  client.js（DSH 客户端插件）",
    [
        "sidebar.panellist「Agent 管理器」入口  ·  main 面板 key=agents",
        "上传卡（拖拽/点击 + 进度条）· 一键转换  ·  Word 实时视图  ·  Agent 日志  ·  下载 DOCX/PDF",
        "流式消费 NDJSON：stage / step / preview / log / done",
    ],
    "#EDF2FE", "#4176E6",
)

box(
    d, 70, 330, 1080, 160,
    "2. Host 插件  index.js（Cordis）",
    [
        "路由 /word-pipeline/*  ·  token + 本机回环 CORS  ·  /convert-stream（NDJSON）",
        "结构理解：analyze(template/source) → LLM 映射 → 校验 → 编排  ·  运行账本 logs/runs/*.jsonl",
        "LLM 调度 /agent/chat（llm.stream）·  Word 实时同步 /open-word  ·  预览渲染队列（单 Word 实例）",
    ],
    "#F3EEFE", "#7C3AED",
)

box(
    d, 70, 540, 1080, 170,
    "3. Python 引擎  dsh_word/",
    [
        "mapping.py（analyze 结构拆分 / apply 映射执行 / verify 映射校验）· fill.py（原子操作）",
        "smartfill.py（无槽位模板：原地替换，保留段落格式）· checks.py + rules.yaml（确定性校验）",
        "render.py（Word COM → PDF/PNG，步骤流）· live.py（live-open / live-sync）· router.py（批量 apply）",
    ],
    "#EAF7F1", "#1B9E6B",
)

box(
    d, 70, 760, 1080, 130,
    "4. Word / 产物",
    [
        "Word COM（可见窗口）：live-preview.docx 逐步同步，改动段落黄色高亮",
        "out.docx / panel-preview.pdf / panel-preview.png  ·  uploads/  ·  logs/runs/<runId>.jsonl",
    ],
    "#FFF6E5", "#D97706",
)

# ---------------- side node: LLM ----------------
box(
    d, 1210, 330, 390, 160,
    "LLM（DSH 默认模型）",
    [
        "llm.stream(provider, model)",
        "只输出 JSON 动作：",
        '{"action":"convert|fill|check|render|none",',
        '"reply":"…"}',
    ],
    "#F0F9FF", "#0891B2",
)

# ---------------- arrows ----------------
arrow(d, 610, 280, 610, 322, "POST /convert-stream  ·  NDJSON 事件")
arrow(d, 610, 490, 610, 532, "python -m dsh_word  ·  stdout PROGRESS / STEP")
arrow(d, 610, 710, 610, 752, "docx / pdf / png  ·  Word COM（步骤同步）")

arrow(d, 1150, 396, 1202, 396)
arrow(d, 1202, 454, 1150, 454)
d.text((1000, 398), "→ llm.stream", font=SMALL, fill="#0891B2")
d.text((1000, 424), "← JSON 动作", font=SMALL, fill="#0891B2")

# request path back up (user action)
d.line([70, 205, 46, 205, 46, 405, 70, 405], fill="#94A3B8", width=3)
d.polygon([(70, 405), (57, 398), (57, 412)], fill="#94A3B8")
d.text((8, 250), "用户操作", font=SMALL, fill="#64748B")

# legend
d.text((70, 930), "关键约定", font=H2, fill=INK)
d.text((70, 968), "· 模板只提供格式：正文原地替换，标题与段落样式保留", font=BODY, fill=MUTED)
d.text((70, 998), "· 内容只来自资料：文本、图片、公式、表格复制，不重写任何句子", font=BODY, fill=MUTED)
d.text((70, 1028), "· 每次运行有 runId：事件写入 logs/runs，UI 实时显示阶段与耗时", font=BODY, fill=MUTED)
d.text((70, 1058), "· Word 是最终呈现：每一步同步进 live-preview.docx，改动处黄色高亮", font=BODY, fill=MUTED)

out = Path("docs/shots/architecture.png")
out.parent.mkdir(parents=True, exist_ok=True)
img.save(out)
print(out)
