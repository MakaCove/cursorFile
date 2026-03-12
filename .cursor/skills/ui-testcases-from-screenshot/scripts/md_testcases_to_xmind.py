#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将测试用例 Markdown 表格（.md）导出为 XMind 文件（.xmind，Zen/2020+ 兼容）。

用法:
- python md_testcases_to_xmind.py <testcases.md> [-o output.xmind] [--group-by 优先级|类型|模块]

说明:
- 输入为包含用例表格的 Markdown 文件（表头需包含 HEADERS）。
- 输出的脑图结构默认按「模块」分组；优先级、前置、步骤等作为用例节点的末级信息展示。
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
import re
import uuid


HEADERS = ["模块", "ID", "标题", "优先级", "类型", "前置条件", "步骤", "测试数据", "预期结果", "备注/覆盖点"]

_RE_MD_TABLE_ROW = re.compile(r"^\s*\|.*\|\s*$")
_RE_MD_TABLE_SEP = re.compile(r"^\s*\|\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|\s*$")


def _new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:12]}"


def _split_md_row(line: str) -> list[str]:
    """
    Split a markdown table row into cells.
    Supports escaped pipe: \\|
    """
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]

    cells: list[str] = []
    cur: list[str] = []
    escape = False
    for ch in s:
        if escape:
            cur.append(ch)
            escape = False
            continue
        if ch == "\\":
            escape = True
            continue
        if ch == "|":
            cells.append("".join(cur).strip())
            cur = []
            continue
        cur.append(ch)
    cells.append("".join(cur).strip())
    return cells


def load_rows_from_md(md_path: Path) -> list[dict]:
    """
    Parse the first markdown table whose header contains required HEADERS.
    Returns list[dict] with keys = HEADERS.
    """
    text = md_path.read_text(encoding="utf-8")
    lines = text.splitlines()

    for i, line in enumerate(lines):
        if not _RE_MD_TABLE_ROW.match(line):
            continue
        header_cells = _split_md_row(line)
        if header_cells and all(h in header_cells for h in HEADERS):
            if i + 1 >= len(lines) or not _RE_MD_TABLE_SEP.match(lines[i + 1]):
                continue
            col_index = {name: header_cells.index(name) for name in HEADERS}
            rows: list[dict] = []
            for j in range(i + 2, len(lines)):
                row_line = lines[j]
                if not row_line.strip():
                    continue
                if not _RE_MD_TABLE_ROW.match(row_line):
                    break
                row_cells = _split_md_row(row_line)
                if len(row_cells) < len(header_cells):
                    row_cells += [""] * (len(header_cells) - len(row_cells))
                row = {h: row_cells[col_index[h]] for h in HEADERS}
                rows.append(row)
            if rows:
                return rows

    raise ValueError("未在 Markdown 中找到包含完整表头的测试用例表格（HEADERS）。")


def _norm(s: str) -> str:
    return (s or "").strip()


def _topic(title: str, children: list[dict] | None = None) -> dict:
    t = {"id": _new_id("topic"), "class": "topic", "title": title}
    if children:
        t["children"] = {"attached": children}
    return t


def _group_key(row: dict, group_by: str) -> str:
    if group_by == "模块":
        return _norm(row.get("模块", "")) or "（未填模块）"
    if group_by == "类型":
        return _norm(row.get("类型", "")) or "（未填类型）"
    # default: 优先级
    return _norm(row.get("优先级", "")) or "（未填优先级）"


def _split_lines(s: str) -> list[str]:
    """
    Split content into lines for XMind child topics.
    Supports literal newlines and <br> (from xlsx_testcases_to_md.py).
    """
    t = _norm(s)
    if not t:
        return []
    t = t.replace("<br>", "\n")
    return [x.strip() for x in t.splitlines() if x.strip()]


def build_xmind_content(rows: list[dict], root_title: str, group_by: str) -> list[dict]:
    """
    Build XMind Zen/2020+ content.json structure (minimal).
    """
    # 分组顺序：
    # - 优先级：按 P0 / P1 / P2 固定顺序
    # - 模块：按在用例表中“首次出现”的顺序（通常与 Excel 顺序一致）
    # - 类型：按字母/汉字排序
    groups: dict[str, list[dict]] = {}
    module_order: list[str] = []
    for r in rows:
        k = _group_key(r, group_by)
        if k not in groups:
            groups[k] = []
            if group_by == "模块":
                module_order.append(k)
        groups[k].append(r)

    group_topics: list[dict] = []
    if group_by == "优先级":
        order = {"P0": 0, "P1": 1, "P2": 2}

        def sort_key(k: str) -> tuple[int, str]:
            return (order.get(k, 9), k)

        group_iter = sorted(groups.keys(), key=sort_key)
    elif group_by == "模块":
        group_iter = module_order
    else:  # 类型等按字母排序
        group_iter = sorted(groups.keys())

    for g in group_iter:
        case_topics: list[dict] = []
        for r in groups[g]:
            tid = _norm(r.get("ID", ""))
            title = _norm(r.get("标题", ""))
            case_title = f"{tid} {title}".strip() if tid or title else "（未命名用例）"

            details: list[dict] = []
            prio = _norm(r.get("优先级", ""))
            typ = _norm(r.get("类型", ""))
            pre = _norm(r.get("前置条件", ""))
            steps = _norm(r.get("步骤", ""))
            data = _norm(r.get("测试数据", ""))
            exp = _norm(r.get("预期结果", ""))
            note = _norm(r.get("备注/覆盖点", ""))

            # Put meta information as leaf nodes under each testcase (not as grouping).
            if prio:
                details.append(_topic(f"优先级：{prio}"))
            if typ:
                details.append(_topic(f"类型：{typ}"))
            if pre:
                details.append(_topic(f"前置：{pre}"))
            if steps:
                step_children = [_topic(x) for x in _split_lines(steps)]
                details.append(_topic("步骤", step_children if step_children else None))
            if data and data != "—":
                details.append(_topic(f"数据：{data}"))
            if exp:
                exp_children = [_topic(x) for x in _split_lines(exp)]
                details.append(_topic("预期", exp_children if exp_children else None))
            if note:
                details.append(_topic(f"备注：{note}"))

            case_topics.append(_topic(case_title, details if details else None))

        group_topics.append(_topic(g, case_topics))

    sheet_id = _new_id("sheet")
    root = _topic(root_title, group_topics)
    # XMind typically requires a structure class on root topic
    root.setdefault("structureClass", "org.xmind.ui.logic.right")

    sheet = {
        "id": sheet_id,
        "class": "sheet",
        "title": "Sheet 1",
        "rootTopic": root,
    }
    return [sheet]


def write_xmind(output_path: Path, content: list[dict]) -> None:
    """
    Write a minimal .xmind (zip) for XMind Zen/2020+.
    """
    created = datetime.now(timezone.utc).isoformat()
    # XMind Zen/2020+ uses metadata.json for workbook info.
    # Keep it minimal but include activeSheetId to avoid "corrupted" warnings.
    active_sheet_id = content[0]["id"] if content and isinstance(content, list) else None
    metadata = {
        "creator": {
            "name": "md_testcases_to_xmind.py",
            "version": "1.0",
        },
        "activeSheetId": active_sheet_id,
        "created": created,
        "modified": created,
    }
    manifest = {
        "file-entries": {
            "content.json": {},
            "metadata.json": {},
            "manifest.json": {},
        }
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("content.json", json.dumps(content, ensure_ascii=False, indent=2))
        z.writestr("metadata.json", json.dumps(metadata, ensure_ascii=False, indent=2))
        z.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))


@dataclass(frozen=True)
class Args:
    md_file: str
    output: str | None
    group_by: str


def main() -> None:
    parser = argparse.ArgumentParser(description="将测试用例 Markdown 表格导出为 XMind")
    parser.add_argument("md_file", help="测试用例 Markdown 文件路径（.md，表格格式）")
    parser.add_argument("-o", "--output", default=None, help="输出 .xmind 路径，默认与 MD 同目录同主名")
    parser.add_argument(
        "--group-by",
        default="模块",
        choices=["优先级", "类型", "模块"],
        help="脑图分组维度（默认：模块）",
    )
    ns = parser.parse_args()
    args = Args(md_file=ns.md_file, output=ns.output, group_by=ns.group_by)

    md_path = Path(args.md_file)
    if not md_path.exists():
        print(f"错误: 文件不存在 {md_path}", file=sys.stderr)
        sys.exit(1)
    if md_path.suffix.lower() != ".md":
        print("错误: 仅支持 .md 输入", file=sys.stderr)
        sys.exit(1)

    try:
        rows = load_rows_from_md(md_path)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)

    root_title = md_path.stem.replace("-测试用例", "") or md_path.stem
    content = build_xmind_content(rows, root_title=root_title, group_by=args.group_by)

    out = Path(args.output) if args.output else md_path.with_suffix(".xmind")
    write_xmind(out, content)
    print(f"已导出: {out}")


if __name__ == "__main__":
    main()

