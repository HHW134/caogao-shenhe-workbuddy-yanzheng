# -*- coding: utf-8 -*-
"""
草案审核 WorkBuddy 验证 —— 命令行入口
=====================================

用法示例：

    # 默认执行「术语与引用联动审核」
    python cli.py 标准草案.docx

    # 三个模块全跑，产物写到 output/
    python cli.py 标准草案.docx --all -o output

    # 只看问题表，不生成 docx/xlsx
    python cli.py 标准草案.docx --all --no-docx --no-xlsx

    # 额外导出 JSON 结果
    python cli.py 标准草案.docx --json output/result.json
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from review_engine import MODULES, run_review  # noqa: E402

_LEVEL_CN = {"error": "错误", "warning": "警告", "info": "提示"}


def _w(text: str) -> int:
    """按中文字符占两列估算显示宽度。"""
    return sum(2 if ord(ch) > 0x2E80 else 1 for ch in str(text))


def _cut(text: str, width: int) -> str:
    text = str(text).replace("\n", " ").strip()
    if _w(text) <= width:
        return text + " " * (width - _w(text))
    out = ""
    for ch in text:
        if _w(out) + _w(ch) > width - 2:
            break
        out += ch
    out += ".."
    return out + " " * max(0, width - _w(out))


def print_table(summary: dict) -> None:
    issues = summary.get("问题列表", [])
    cols = [("序号", 5), ("审核要点", 24), ("级别", 6), ("问题", 46), ("位置", 30)]
    line = "+".join("-" * (w + 2) for _, w in cols)
    print("+" + line + "+")
    print("| " + " | ".join(_cut(h, w) for h, w in cols) + " |")
    print("+" + line + "+")
    for rec in issues:
        row = [
            rec.get("序号", ""), rec.get("审核要点", ""),
            _LEVEL_CN.get(rec.get("级别"), rec.get("级别", "")),
            rec.get("问题", ""), rec.get("位置", ""),
        ]
        print("| " + " | ".join(_cut(v, w) for v, (_, w) in zip(row, cols)) + " |")
    if not issues:
        print("| " + _cut("未发现问题", sum(w for _, w in cols) + 3 * (len(cols) - 1)) + " |")
    print("+" + line + "+")


def print_suggestions(summary: dict, limit: int = 20) -> None:
    issues = summary.get("问题列表", [])
    if not issues:
        return
    print("\n【修改建议】")
    for rec in issues[:limit]:
        print(f"  {rec.get('序号')}. [{_LEVEL_CN.get(rec.get('级别'), '')}] "
              f"{rec.get('审核要点')}：{rec.get('修改建议')}")
    if len(issues) > limit:
        print(f"  …… 其余 {len(issues) - limit} 条见 xlsx 清单")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="草案审核 WorkBuddy 验证：标准草案 docx 自动审核（批注 + 问题清单）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="可用模块：\n" + "\n".join(
            f"  {k:<10} {v['name']} —— {v['desc']}" for k, v in MODULES.items()),
    )
    parser.add_argument("input", help="待审核的 .docx 文件")
    parser.add_argument("-m", "--modules", default="chapters",
                        help="模块，逗号分隔：chapters,normative,fulldoc（默认 chapters）")
    parser.add_argument("--all", action="store_true", help="执行全部模块")
    parser.add_argument("-o", "--out-dir", default=None, help="产物输出目录（默认与源文件同目录）")
    parser.add_argument("--json", dest="json_path", default=None, help="额外导出 JSON 结果")
    parser.add_argument("--no-docx", action="store_true", help="不生成带批注的 docx")
    parser.add_argument("--no-xlsx", action="store_true", help="不生成问题清单 xlsx")
    args = parser.parse_args(argv)

    modules = list(MODULES) if args.all else [m.strip() for m in args.modules.split(",") if m.strip()]

    try:
        summary = run_review(
            args.input, modules=modules, out_dir=args.out_dir,
            write_docx=not args.no_docx, write_xlsx=not args.no_xlsx,
        )
    except FileNotFoundError as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2

    print(f"\n文件：{summary['文件名']}")
    print(f"模块：{'、'.join(summary['执行模块'])}")
    stats = summary["级别统计"]
    print(f"结论：{summary['审核结论']}　问题 {summary['问题总数']} 条"
          f"（错误 {stats.get('error', 0)} / 警告 {stats.get('warning', 0)} / 提示 {stats.get('info', 0)}）\n")

    print_table(summary)
    print_suggestions(summary)

    if summary.get("模块异常"):
        print("\n【模块异常】")
        for mod, err in summary["模块异常"].items():
            print(f"  {mod}: {err}")

    if summary.get("产物"):
        print("\n【产物】")
        for art in summary["产物"]:
            print(f"  {art['类型']}（{art['模块']}）：{art['路径']}")

    if args.json_path:
        Path(args.json_path).parent.mkdir(parents=True, exist_ok=True)
        with open(args.json_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"\nJSON 结果：{args.json_path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
