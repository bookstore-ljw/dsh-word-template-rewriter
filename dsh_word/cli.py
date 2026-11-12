"""Command line surface: ``fill`` and ``check`` are the two worker entry points."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .checks import load_rules, run_checks
from .fill import fill_docx, inspect_docx, load_plan

_SEVERITY_ORDER = {"info": 0, "warning": 1, "error": 2}


def _cmd_fill(args: argparse.Namespace) -> int:
    emit = (
        (lambda item: print("PROGRESS " + json.dumps(item, ensure_ascii=False), flush=True))
        if args.progress
        else None
    )
    if args.auto_plan:
        if not args.source:
            print("--auto-plan 需要同时给出 --source", file=sys.stderr)
            return 2
        from .autoplan import build_auto_plan
        from .smartfill import has_markers, smart_fill

        if not has_markers(args.template):
            smart = smart_fill(
                args.template,
                args.source,
                args.out,
                steps_dir=args.steps_dir,
                on_op=emit,
            )
            print(f"已生成: {smart.out_path}")
            print(
                f"模板没有 {{{{ }}}} 槽位，已按结构自动填充："
                f"{smart.markers} 处，追加段落 {smart.appended}，复制表格 {smart.tables}"
            )
            for item in smart.mapping:
                print(f"  {item['kind']:<8} {item['slot']:<20} ← {item['from']}")
            return 0

        asset_dir = Path(args.asset_dir) if args.asset_dir else Path(args.out).parent / "assets"
        auto = build_auto_plan(args.template, args.source, asset_dir=asset_dir)
        plan = auto.plan
        for item in auto.mapping:
            print(f"  映射 {item['slot']:<24} ← {item['from']} ({item['kind']})")
        if auto.unresolved:
            print("自动映射未覆盖以下槽位：", file=sys.stderr)
            for slot in auto.unresolved:
                print(f"  {slot}", file=sys.stderr)
            return 2
    else:
        if not args.plan:
            print("需要 --plan，或使用 --auto-plan --source", file=sys.stderr)
            return 2
        plan = load_plan(args.plan)
    result = fill_docx(
        args.template,
        plan,
        args.out,
        steps_dir=args.steps_dir,
        trace_path=args.trace,
        on_op=(
            (lambda item: print("PROGRESS " + json.dumps(item, ensure_ascii=False), flush=True))
            if args.progress
            else None
        ),
    )
    if not result.trace:
        print("模板中没有找到 {{ }} 槽位标记，未做任何填充。", file=sys.stderr)
        print(
            "请确认资料/模板没有传反，并在模板里放好占位符，例如 {{ text.title }}。",
            file=sys.stderr,
        )
        return 2
    print(f"已生成: {result.out_path}")
    print(f"原子操作: {len(result.trace)}")
    for item in result.trace:
        print(f"  {item['op']:02d} {item['kind']:<8} {item['slot']:<24} {item['detail']}")
    if result.steps:
        print(f"中间版本: {len(result.steps)} 个，位于 {args.steps_dir}")
    if result.remaining:
        print("未替换占位符:", file=sys.stderr)
        for item in result.remaining:
            print(f"  {item['slot']} @ {item['location']}", file=sys.stderr)
        return 2
    return 0


def _cmd_check(args: argparse.Namespace) -> int:
    findings = run_checks(args.docx, rules_path=args.rules, template_path=args.template)
    counts = {"error": 0, "warning": 0, "info": 0}
    for finding in findings:
        counts[finding.severity] = counts.get(finding.severity, 0) + 1
        print(f"[{finding.severity}] {finding.rule} @ {finding.location}: {finding.message}")
        if finding.snippet:
            print(f"    {finding.snippet}")
        if finding.detail:
            print(f"    {finding.detail}")
    print(
        f"合计: error={counts.get('error', 0)} warning={counts.get('warning', 0)} "
        f"info={counts.get('info', 0)}"
    )
    if args.json:
        Path(args.json).write_text(
            json.dumps([finding.as_dict() for finding in findings], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"报告: {args.json}")
    threshold = _SEVERITY_ORDER[args.fail_on]
    worst = max((_SEVERITY_ORDER[f.severity] for f in findings), default=-1)
    return 1 if worst >= threshold else 0


def _cmd_inspect(args: argparse.Namespace) -> int:
    print(json.dumps(inspect_docx(args.docx), ensure_ascii=False, indent=2))
    return 0


def _cmd_apply(args: argparse.Namespace) -> int:
    from .router import run_router

    summary = run_router(
        args.template,
        args.base_plan,
        args.corrections,
        args.out_dir,
        report_path=args.json,
        dry_run=args.dry_run,
    )
    if args.dry_run:
        print(f"待处理修正 {summary['corrections']} 条 → {len(summary['batches'])} 个批次")
        for index, batch in enumerate(summary["batches"], start=1):
            print(f"  {index:02d} {batch['kind']:<7} {', '.join(batch['slots'])}")
        return 0
    for index, result in enumerate(summary["results"], start=1):
        errors = sum(1 for item in result["findings"] if item["severity"] == "error")
        warnings = sum(1 for item in result["findings"] if item["severity"] == "warning")
        print(
            f"批次 {index:02d} {result['kind']:<7} → {result['output']} "
            f"(操作 {result['operations']}，error={errors} warning={warnings})"
        )
        for item in result["findings"]:
            print(f"    [{item['severity']}] {item['rule']} @ {item['location']}")
        for item in result["remaining"]:
            print(f"    未替换 {item['slot']} @ {item['location']}")
    print(f"整体: {'通过' if summary.get('ok') else '未通过'}")
    if args.json:
        print(f"报告: {args.json}")
    return 0 if summary.get("ok") else 1


def _cmd_analyze(args: argparse.Namespace) -> int:
    from .mapping import analyze

    print(json.dumps(analyze(args.docx), ensure_ascii=False))
    return 0


def _cmd_map_verify(args: argparse.Namespace) -> int:
    from .mapping import analyze, verify_mapping

    mapping = json.loads(Path(args.mapping).read_text(encoding="utf-8-sig"))
    result = verify_mapping(
        analyze(args.template),
        analyze(args.source),
        mapping.get("mapping", mapping) if isinstance(mapping, dict) else mapping,
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["ok"] else 2


def _cmd_map_apply(args: argparse.Namespace) -> int:
    from .mapping import apply_mapping

    payload = json.loads(Path(args.mapping).read_text(encoding="utf-8-sig"))
    mapping = payload.get("mapping", []) if isinstance(payload, dict) else payload
    result = apply_mapping(
        args.template,
        args.source,
        mapping,
        args.out,
        steps_dir=args.steps_dir,
        on_op=(
            (lambda item: print("PROGRESS " + json.dumps(item, ensure_ascii=False), flush=True))
            if args.progress
            else None
        ),
        on_check=(
            (lambda item: print("CHECK " + json.dumps(item, ensure_ascii=False), flush=True))
            if args.progress
            else None
        ),
    )
    print(json.dumps({"out": result.out_path, "applied": result.applied, "skipped": result.skipped}, ensure_ascii=False))
    return 0


def _cmd_heal(args: argparse.Namespace) -> int:
    from .heal import heal_format

    cost = heal_format(args.template, args.docx)
    print(json.dumps(cost, ensure_ascii=False))
    return 0


def _cmd_safe_fill(args: argparse.Namespace) -> int:
    from .safefill import safe_fill

    def emit(item: dict) -> None:
        print("SEGMENT " + json.dumps(item, ensure_ascii=False), flush=True)

    outcome = safe_fill(
        args.template,
        args.source,
        args.out,
        steps_dir=args.steps_dir,
        on_event=emit if args.progress else None,
    )
    print(json.dumps({"out": outcome["out"], "reports": outcome["reports"],
                      "leftover": len(outcome["leftover"])}, ensure_ascii=False))
    return 0


def _cmd_sectional(args: argparse.Namespace) -> int:
    from .sectional import sectional_fill

    def emit(item: dict) -> None:
        print("SEGMENT " + json.dumps(item, ensure_ascii=False), flush=True)

    outcome = sectional_fill(
        args.template,
        args.source,
        args.out,
        steps_dir=args.steps_dir,
        on_event=emit if args.progress else None,
    )
    print(json.dumps({"out": outcome["out"], "reports": outcome["reports"]}, ensure_ascii=False))
    return 0


def _cmd_compose(args: argparse.Namespace) -> int:
    from .compose import compose

    emit = (
        (lambda item: print("PROGRESS " + json.dumps(item, ensure_ascii=False), flush=True))
        if args.progress
        else None
    )
    result = compose(args.template, args.source, args.out, steps_dir=args.steps_dir, on_op=emit)
    print(json.dumps({"out": result.out_path, "headings": result.headings, "paragraphs": result.paragraphs,
                      "equations": result.equations, "figures": result.figures, "tables": result.tables}, ensure_ascii=False))
    return 0


def _cmd_section_fill(args: argparse.Namespace) -> int:
    from .sectionfill import section_fill

    def emit(event: dict) -> None:
        kind = event.pop("type")
        print(kind.upper() + " " + json.dumps(event, ensure_ascii=False), flush=True)

    outcome = section_fill(
        args.template,
        args.source,
        args.out,
        steps_dir=args.steps_dir,
        on_event=emit if args.progress else None,
    )
    payload = {"out": args.out}
    if outcome["missing"]:
        payload["missing"] = outcome["missing"]
    print(json.dumps(payload, ensure_ascii=False))
    return 0


def _cmd_live_open(args: argparse.Namespace) -> int:
    from .live import open_file

    print(open_file(args.file))
    return 0


def _cmd_live_sync(args: argparse.Namespace) -> int:
    from .live import sync

    print(sync(args.target, args.source, save=args.save, focus=args.focus))
    return 0


def _cmd_live_sync_block(args: argparse.Namespace) -> int:
    from .live import sync_block

    print(sync_block(args.target, args.step, args.index, save=args.save))
    return 0


def _cmd_render_steps(args: argparse.Namespace) -> int:
    from .render import render_many

    steps_dir = Path(args.dir)
    paths = sorted(steps_dir.glob(args.pattern))
    if not paths:
        print(f"{steps_dir} 下没有匹配 {args.pattern} 的文件", file=sys.stderr)
        return 2
    manifest = render_many(
        paths,
        args.out_dir,
        dpi=args.dpi,
        on_step=(
            (lambda source, png: print(f"STEP {png}", flush=True)) if args.stream else None
        ),
    )
    payload = json.dumps(manifest, ensure_ascii=False, indent=2)
    print(payload)
    if args.json:
        Path(args.json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json).write_text(payload, encoding="utf-8")
    return 0


def _cmd_retypeset(args: argparse.Namespace) -> int:
    from .retypeset import Profile, retypeset, verify_layout

    profile = Profile.load(args.profile)
    overrides = {}
    if args.figure_width_cm is not None:
        overrides["figure_width_cm"] = args.figure_width_cm
    if args.table_width_twip is not None:
        overrides["table_width_twip"] = args.table_width_twip
    if args.table_font_pt is not None:
        overrides["table_font_pt"] = args.table_font_pt
    profile.layout.update(overrides)
    if args.no_keep_next:
        profile.flags["keep_next"] = False

    emit = (
        (lambda item: print("PROGRESS " + json.dumps(item, ensure_ascii=False), flush=True))
        if args.progress
        else None
    )
    result = retypeset(args.template, args.source, args.out, profile, on_event=emit)
    payload = result.to_dict()

    if args.verify:
        pdf = args.pdf or str(Path(args.out).with_suffix(".pdf"))
        payload["layout"] = verify_layout(args.out, pdf)
        payload["pdf"] = pdf

    if args.json:
        Path(args.json).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False))
    issues = (payload.get("layout") or {}).get("issues") or []
    return 1 if issues else 0


def _cmd_render(args: argparse.Namespace) -> int:
    from .render import export_pdf, render_png

    pdf_path = args.pdf or str(Path(args.docx).with_suffix(".pdf"))
    path, pages = export_pdf(args.docx, pdf_path)
    print(f"PDF: {path} ({pages} 页)")
    if args.png:
        png, _ = render_png(pdf_path, args.png, page=args.page, dpi=args.dpi)
        print(f"PNG: {png}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="dsh-word", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    fill = sub.add_parser("fill", help="把资料按 plan 填充进模板")
    fill.add_argument("--template", required=True)
    fill.add_argument("--plan", help="槽位计划 YAML；与 --auto-plan 二选一")
    fill.add_argument("--source", help="资料 DOCX，供 --auto-plan 使用")
    fill.add_argument("--auto-plan", action="store_true", help="按资料顺序自动生成槽位映射")
    fill.add_argument("--asset-dir", help="自动映射提取图片的落盘目录")
    fill.add_argument("--out", required=True)
    fill.add_argument("--steps-dir", help="每完成一个原子操作写一个中间版本")
    fill.add_argument("--trace", help="写出操作审计 JSON")
    fill.add_argument("--progress", action="store_true", help="逐步输出 PROGRESS JSON 行")
    fill.set_defaults(func=_cmd_fill)

    check = sub.add_parser("check", help="对已生成文档执行确定性检查")
    check.add_argument("--docx", required=True)
    check.add_argument("--rules", help="检查规则 YAML，默认使用内置 rules.yaml")
    check.add_argument("--template", help="模板路径，用于格式/表格一致性比对")
    check.add_argument("--json", help="写出 JSON 报告")
    check.add_argument(
        "--fail-on", choices=("none", "info", "warning", "error"), default="error"
    )
    check.set_defaults(func=_cmd_check)

    inspect = sub.add_parser("inspect", help="输出文档结构摘要")
    inspect.add_argument("--docx", required=True)
    inspect.set_defaults(func=_cmd_inspect)

    apply = sub.add_parser("apply", help="Router：合批修正并逐批重填 + 复检")
    apply.add_argument("--template", required=True)
    apply.add_argument("--base-plan", required=True)
    apply.add_argument("--corrections", required=True)
    apply.add_argument("--out-dir", required=True)
    apply.add_argument("--json", help="写出批次与检查报告")
    apply.add_argument("--dry-run", action="store_true", help="只显示合批结果，不生成文档")
    apply.set_defaults(func=_cmd_apply)

    steps = sub.add_parser("render-steps", help="一次 Word 实例批量渲染步骤文档")
    steps.add_argument("--dir", required=True)
    steps.add_argument("--out-dir", required=True)
    steps.add_argument("--pattern", default="*.docx")
    steps.add_argument("--dpi", type=int, default=80)
    steps.add_argument("--json")
    steps.add_argument("--stream", action="store_true")
    steps.set_defaults(func=_cmd_render_steps)

    analyze_p = sub.add_parser("analyze", help="分析文档结构（JSON）")
    analyze_p.add_argument("--docx", required=True)
    analyze_p.set_defaults(func=_cmd_analyze)

    map_verify = sub.add_parser("map-verify", help="校验结构映射是否合适")
    map_verify.add_argument("--template", required=True)
    map_verify.add_argument("--source", required=True)
    map_verify.add_argument("--mapping", required=True)
    map_verify.set_defaults(func=_cmd_map_verify)

    map_apply = sub.add_parser("map-apply", help="按结构映射填充模板")
    map_apply.add_argument("--template", required=True)
    map_apply.add_argument("--source", required=True)
    map_apply.add_argument("--mapping", required=True)
    map_apply.add_argument("--out", required=True)
    map_apply.add_argument("--steps-dir")
    map_apply.add_argument("--progress", action="store_true")
    map_apply.set_defaults(func=_cmd_map_apply)

    heal = sub.add_parser("heal", help="把模板格式恢复到输出文档")
    heal.add_argument("--template", required=True)
    heal.add_argument("--docx", required=True)
    heal.set_defaults(func=_cmd_heal)

    safe = sub.add_parser("safe-fill", help="非破坏填充：只改已有段落文字，保留模板版式")
    safe.add_argument("--template", required=True)
    safe.add_argument("--source", required=True)
    safe.add_argument("--out", required=True)
    safe.add_argument("--steps-dir")
    safe.add_argument("--progress", action="store_true")
    safe.set_defaults(func=_cmd_safe_fill)

    sectional = sub.add_parser("sectional", help="按模板章节逐段原地填充（不重建全文）")
    sectional.add_argument("--template", required=True)
    sectional.add_argument("--source", required=True)
    sectional.add_argument("--out", required=True)
    sectional.add_argument("--steps-dir")
    sectional.add_argument("--progress", action="store_true")
    sectional.set_defaults(func=_cmd_sectional)

    compose_p = sub.add_parser("compose", help="按模板样式重建文档（标题/段落/公式/图/表编号）")
    compose_p.add_argument("--template", required=True)
    compose_p.add_argument("--source", required=True)
    compose_p.add_argument("--out", required=True)
    compose_p.add_argument("--steps-dir")
    compose_p.add_argument("--progress", action="store_true")
    compose_p.set_defaults(func=_cmd_compose)

    retype = sub.add_parser(
        "retypeset",
        help="按模板重排整篇文档（跨模板重建：换样式/编号/题注，保留公式·图·表·run 级格式）",
    )
    retype.add_argument("--template", required=True)
    retype.add_argument("--source", required=True)
    retype.add_argument("--out", required=True)
    retype.add_argument("--profile", help="profile 名称或路径，默认 ieee-to-ttp")
    retype.add_argument("--figure-width-cm", type=float)
    retype.add_argument("--table-width-twip", type=int)
    retype.add_argument("--table-font-pt", type=int)
    retype.add_argument("--no-keep-next", action="store_true", help="不给表格/图题加「与下段同页」")
    retype.add_argument("--verify", action="store_true", help="用 Word 复核分页并导出 PDF")
    retype.add_argument("--pdf")
    retype.add_argument("--json")
    retype.add_argument("--progress", action="store_true")
    retype.set_defaults(func=_cmd_retypeset)

    section = sub.add_parser("section-fill", help="按模板章节逐段填充")
    section.add_argument("--template", required=True)
    section.add_argument("--source", required=True)
    section.add_argument("--out", required=True)
    section.add_argument("--steps-dir")
    section.add_argument("--progress", action="store_true")
    section.set_defaults(func=_cmd_section_fill)

    live_open = sub.add_parser("live-open", help="在 Word 中打开文档（实时查看）")
    live_open.add_argument("--file", required=True)
    live_open.set_defaults(func=_cmd_live_open)

    live_sync = sub.add_parser("live-sync", help="把步骤文档同步进已打开的 Word 窗口")
    live_sync.add_argument("--target", required=True)
    live_sync.add_argument("--source", required=True)
    live_sync.add_argument("--no-save", dest="save", action="store_false", default=True)
    live_sync.add_argument("--focus", type=int)
    live_sync.set_defaults(func=_cmd_live_sync)

    live_block = sub.add_parser("live-sync-block", help="只更新一段并跳转高亮")
    live_block.add_argument("--target", required=True)
    live_block.add_argument("--step", required=True)
    live_block.add_argument("--index", type=int, required=True)
    live_block.add_argument("--no-save", dest="save", action="store_false", default=True)
    live_block.set_defaults(func=_cmd_live_sync_block)

    render = sub.add_parser("render", help="用 Word 渲染为 PDF/PNG 做版面抽检")
    render.add_argument("--docx", required=True)
    render.add_argument("--pdf")
    render.add_argument("--png")
    render.add_argument("--page", type=int, default=0)
    render.add_argument("--dpi", type=int, default=120)
    render.set_defaults(func=_cmd_render)

    rules = sub.add_parser("rules", help="列出已启用的检查规则")
    rules.set_defaults(func=lambda args: _print_rules())
    return parser


def _print_rules() -> int:
    for rule in load_rules():
        print(f"{rule['id']:<22} {rule['kind']:<18} {rule.get('severity', 'warning')}")
    return 0


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8")
            except (OSError, ValueError):
                pass
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
