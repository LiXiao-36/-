#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
龙洲校区「家作一周一上传」催交名单生成器（命令行版）

用法：
  python cli.py                 # 默认扫描最新一周
  python cli.py --week 7        # 指定第 7 周
  python cli.py --all           # 扫描全部周，并导出总览矩阵
  python cli.py --list          # 只列出各周目录
  python cli.py --dir "D:\\某周目录"

图形界面请运行 gui.py（或双击 打开界面.bat）。
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import checker
from checker import (
    DEFAULT_ROSTER_XLS, DEFAULT_YEAR_DIR, GRADE_ORDER, OUT_DIR,
    analyze_week, analyze_weeks, count_files, fmt_time, list_weeks,
    load_roster, pad, setup_console_hint, write_csv, write_matrix_xlsx,
    write_messages, write_xlsx,
)


def resolve_week_dir(year_dir: Path, week: int | None, explicit: str | None):
    if explicit:
        p = Path(explicit)
        if not p.is_dir():
            raise FileNotFoundError(f"目录不存在：{p}")
        import re
        m = checker.WEEK_DIR_RE.search(p.name)
        return (int(m.group(1)) if m else None), p

    weeks = list_weeks(year_dir)
    if not weeks:
        raise FileNotFoundError(f"{year_dir} 下没有找到「第N周」目录")
    if week is None:
        return weeks[-1]
    for n, p in weeks:
        if n == week:
            return n, p
    raise FileNotFoundError(
        f"没有第 {week} 周目录。现有：" + "、".join(f"第{n}周" for n, _ in weeks)
    )


def print_week_report(res, roster) -> None:
    print("=" * 62)
    print(f"  {res.label}   家作上交情况")
    print(f"  目录：{res.path}")
    print(f"  时间：{fmt_time(res.scanned_at)}")
    print("=" * 62)
    print(f"  班级总数 {res.total}    已交 {res.ok_count}    未交 {res.missing_count}")
    print()

    if res.missing:
        print("  【未交班级】")
        by_grade: dict[str, list[dict]] = {}
        for row in res.missing:
            by_grade.setdefault(row["grade"], []).append(row)
        for g, rows in by_grade.items():
            names = "、".join(f"{r['code']} {r['teacher'] or '?'}" for r in rows)
            print(f"    {pad(g, 8)}（{len(rows)}个）：{names}")
        print()
        print("  【需要联系的班主任】")
        for i, row in enumerate(res.missing, 1):
            print(f"    {i:>2}. {pad(row['teacher'] or '（配班表无姓名）', 10)}"
                  f"{row['code']}班   配班：{row['partner'] or '—'}")
        print()
    else:
        print("  ✅ 全部班级都已上交。")
        print()

    if res.submitted_rows:
        print("  【已交班级】")
        line = "、".join(f"{r['code']}({len(r['files'])})" for r in res.submitted_rows)
        print(f"    {line}")
        print()

    if res.zero_grades:
        print(f"  ⚠ 注意：{'、'.join(res.zero_grades)}整年级 0 个文件。")
        print("    这可能不是个别班漏交，而是还没到上传时间，或文件交到了别的目录/别的周。")
        print("    建议先核实，再决定要不要群发催交。")
        print()

    if res.unknown:
        print(f"  ⚠ 有 {len(res.unknown)} 个文件无法识别班号：")
        for u in res.unknown[:8]:
            print(f"    - {u}")
        if len(res.unknown) > 8:
            print(f"    …等 {len(res.unknown)} 个")
        print()


def main() -> int:
    setup_console_hint()
    ap = argparse.ArgumentParser(
        description="统计家作上交情况，生成未交班级的班主任催交名单",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("--week", type=int, help="要统计的周次，如 7；默认取最新一周")
    ap.add_argument("--all", action="store_true", help="扫描全部周，并导出总览矩阵")
    ap.add_argument("--dir", help="直接指定某一周目录（优先级最高）")
    ap.add_argument("--year-dir", default=str(DEFAULT_YEAR_DIR), help="学年根目录")
    ap.add_argument("--roster", default=str(DEFAULT_ROSTER_XLS), help="配班表路径")
    ap.add_argument("--list", action="store_true", help="只列出各周目录")
    ap.add_argument("--refresh-roster", action="store_true", help="强制重新读取配班表")
    args = ap.parse_args()

    year_dir = Path(args.year_dir)
    roster_path = Path(args.roster)

    if args.list:
        weeks = list_weeks(year_dir)
        print(f"{year_dir}\n")
        for n, p in weeks:
            print(f"  第{n}周  {p.name}    文件数 {count_files(p)}")
        return 0

    roster = load_roster(roster_path, force=args.refresh_roster,
                         log=lambda m: print(f"[配班表] {m}"))
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    if args.all:
        pairs = list_weeks(year_dir)
        results = analyze_weeks(roster, pairs)
        for res in results:
            print()
            print_week_report(res, roster)
        path = OUT_DIR / f"上交总览_{year_dir.name}_{stamp}.xlsx"
        write_matrix_xlsx(path, results, roster, year_dir)
        print("=" * 62)
        print(f"  已导出总览矩阵：{path}")
        print("=" * 62)
        return 0

    week_no, week_dir = resolve_week_dir(year_dir, args.week, args.dir)
    res = analyze_week(roster, week_no, week_dir)
    print()
    print_week_report(res, roster)

    tag = f"第{week_no}周" if week_no is not None else week_dir.name
    xlsx_path = OUT_DIR / f"催交名单_{tag}_{stamp}.xlsx"
    csv_path = OUT_DIR / f"催交名单_{tag}_{stamp}.csv"
    write_xlsx(xlsx_path, res, year_dir)
    write_csv(csv_path, res.missing)
    msg_single, msg_group = write_messages(OUT_DIR, tag, res, year_dir)

    print("  【已生成】")
    print(f"    名单表格：{xlsx_path.name}")
    print(f"    名单 CSV：{csv_path.name}")
    print(f"    逐个发送文案：{msg_single.name}")
    print(f"    群发文案：{msg_group.name}")
    print(f"    所在目录：{OUT_DIR}")
    print("=" * 62)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n已取消")
        sys.exit(130)
    except Exception as exc:
        print(f"\n[错误] {exc}", file=sys.stderr)
        sys.exit(1)
