#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将测试用例 Markdown 表格（.md）导出为 XMind 文件（.xmind，Zen/2020+ 兼容）。

支持两类用例表：
  - 功能点用例（表头含"模块""ID""标题"）→ 按模块/优先级/类型分组
  - 业务链路用例（表头含"链路ID""链路名称""涉及模块"）→ 按链路类型分组

用法:
  python md_testcases_to_xmind.py <testcases.md> [testcases2.md ...] [-o output.xmind] [--group-by 模块|优先级|类型]

说明:
  - 输入为一个或多个包含用例表格的 Markdown 文件。
  - 多个文件的内容会合并到同一个 XMind 中，功能点用例与业务链路用例分别放在两个一级分支下。
  - 支持将分开的「功能点测试用例.md」和「业务链路测试用例.md」合并为一个 XMind。
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
from typing import Optional


FUNCTIONAL_HEADERS = ["模块", "ID", "标题", "优先级", "类型", "前置条件", "步骤", "测试数据", "预期结果", "备注/覆盖点"]
CHAIN_HEADERS = ["链路ID", "链路名称", "涉及模块", "优先级", "类型", "前置条件", "步骤", "测试数据", "预期结果", "备注/覆盖点"]

_RE_MD_TABLE_ROW = re.compile(r"^\s*\|.*\|\s*$")
_RE_MD_TABLE_SEP = re.compile(r"^\s*\|\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|\s*$")


def _new_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:12]}"


def _split_md_row(line: str) -> list[str]:
    """Split a markdown table row into cells. Supports escaped pipe: \\|"""
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


def _parse_table(lines: list[str], start: int, headers: list[str]) -> tuple[Optional[list[dict]], int]:
    """尝试从 start 行开始解析一个表格。返回 (rows, consumed_lines)。"""
    if start >= len(lines):
        return None, 0

    for i in range(start, len(lines)):
        line = lines[i]
        if not _RE_MD_TABLE_ROW.match(line):
            continue
        header_cells = _split_md_row(line)
        if not header_cells:
            continue
        if all(h in header_cells for h in headers):
            if i + 1 >= len(lines) or not _RE_MD_TABLE_SEP.match(lines[i + 1]):
                continue

            col_index = {name: header_cells.index(name) for name in headers}
            rows: list[dict] = []
            for j in range(i + 2, len(lines)):
                row_line = lines[j]
                if not row_line.strip():
                    continue
                if not _RE_MD_TABLE_ROW.match(row_line):
                    return rows, j - start
                row_cells = _split_md_row(row_line)
                if len(row_cells) < len(header_cells):
                    row_cells += [""] * (len(header_cells) - len(row_cells))
                row = {h: row_cells[col_index[h]] for h in headers}
                rows.append(row)
            if rows:
                return rows, j - start + 1
            return None, 0
    return None, 0


def load_tables_from_md(md_path: Path) -> tuple[list[dict], list[dict]]:
    """从 Markdown 中分别解析功能点用例表和业务链路用例表。"""
    text = md_path.read_text(encoding="utf-8")
    lines = text.splitlines()

    functional_rows: list[dict] = []
    chain_rows: list[dict] = []

    pos = 0
    while pos < len(lines):
        rows, consumed = _parse_table(lines, pos, FUNCTIONAL_HEADERS)
        if rows:
            functional_rows = rows
            pos += max(consumed, 1)
            continue

        rows, consumed = _parse_table(lines, pos, CHAIN_HEADERS)
        if rows:
            chain_rows = rows
            pos += max(consumed, 1)
            continue

        pos += 1

    if not functional_rows and not chain_rows:
        raise ValueError("未在 Markdown 中找到包含完整表头的测试用例表格。")

    return functional_rows, chain_rows


def _norm(s: str) -> str:
    return (s or "").strip()


def _topic(title: str, children: list[dict] | None = None) -> dict:
    t = {"id": _new_id("topic"), "class": "topic", "title": title}
    if children:
        t["children"] = {"attached": children}
    return t


def _group_key_functional(row: dict, group_by: str) -> str:
    if group_by == "模块":
        return _norm(row.get("模块", "")) or "（未填模块）"
    if group_by == "类型":
        return _norm(row.get("类型", "")) or "（未填类型）"
    return _norm(row.get("优先级", "")) or "（未填优先级）"


def _split_lines(s: str) -> list[str]:
    """Split content into lines for XMind child topics. Supports literal newlines and <br>."""
    t = _norm(s)
    if not t:
        return []
    t = t.replace("<br>", "\n")
    return [x.strip() for x in t.splitlines() if x.strip()]


def _build_case_topics(r: dict, is_chain: bool = False) -> tuple[list[dict], str]:
    """Build the detail topic nodes for a single test case. Returns (details, case_title)."""
    if is_chain:
        tid = _norm(r.get("链路ID", ""))
        title = _norm(r.get("链路名称", ""))
        case_title = f"{tid} {title}".strip() if tid or title else "（未命名链路）"
    else:
        tid = _norm(r.get("ID", ""))
        title = _norm(r.get("标题", ""))
        case_title = f"{tid} {title}".strip() if tid or title else "（未命名用例）"

    details: list[dict] = []

    if is_chain:
        modules_val = _norm(r.get("涉及模块", ""))
        if modules_val:
            details.append(_topic(f"涉及模块：{modules_val}"))

    prio = _norm(r.get("优先级", ""))
    typ = _norm(r.get("类型", ""))
    pre = _norm(r.get("前置条件", ""))
    steps = _norm(r.get("步骤", ""))
    data = _norm(r.get("测试数据", ""))
    exp = _norm(r.get("预期结果", ""))
    note = _norm(r.get("备注/覆盖点", ""))

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

    return details, case_title


def build_functional_branch(rows: list[dict], group_by: str) -> list[dict]:
    """Build XMind branch for functional test cases, grouped by selected dimension."""
    groups: dict[str, list[dict]] = {}
    module_order: list[str] = []
    for r in rows:
        k = _group_key_functional(r, group_by)
        if k not in groups:
            groups[k] = []
            if group_by == "模块":
                module_order.append(k)
        groups[k].append(r)

    group_topics: list[dict] = []
    if group_by == "优先级":
        order = {"P0": 0, "P1": 1, "P2": 2}
        group_iter = sorted(groups.keys(), key=lambda k: (order.get(k, 9), k))
    elif group_by == "模块":
        group_iter = module_order
    else:
        group_iter = sorted(groups.keys())

    for g in group_iter:
        case_topics: list[dict] = []
        for r in groups[g]:
            details, case_title = _build_case_topics(r, is_chain=False)
            case_topics.append(_topic(case_title, details if details else None))
        group_topics.append(_topic(g, case_topics))

    return group_topics


def build_chain_branch(rows: list[dict]) -> list[dict]:
    """Build XMind branch for business chain test cases, grouped by chain type (from 备注/覆盖点)."""
    groups: dict[str, list[dict]] = {}
    for r in rows:
        note = _norm(r.get("备注/覆盖点", ""))
        chain_type = "（未分类链路）"
        if "主链路" in note:
            chain_type = "主链路"
        elif "分支链路" in note:
            chain_type = "分支链路"
        elif "中断恢复链路" in note or "中断恢复" in note:
            chain_type = "中断恢复链路"
        elif "回退链路" in note:
            chain_type = "回退链路"
        elif "超时中断" in note:
            chain_type = "超时中断链路"

        if chain_type not in groups:
            groups[chain_type] = []
        groups[chain_type].append(r)

    group_topics: list[dict] = []
    type_order = {"主链路": 0, "分支链路": 1, "中断恢复链路": 2, "超时中断链路": 3, "回退链路": 4}

    for g in sorted(groups.keys(), key=lambda k: type_order.get(k, 9)):
        case_topics: list[dict] = []
        for r in groups[g]:
            details, case_title = _build_case_topics(r, is_chain=True)
            case_topics.append(_topic(case_title, details if details else None))
        group_topics.append(_topic(g, case_topics))

    return group_topics


def build_xmind_content(
    functional_rows: list[dict],
    chain_rows: list[dict],
    root_title: str,
    group_by: str,
) -> list[dict]:
    """Build XMind Zen/2020+ content.json structure (minimal)."""
    root_children: list[dict] = []

    if functional_rows:
        func_children = build_functional_branch(functional_rows, group_by)
        root_children.append(_topic("功能点用例", func_children))

    if chain_rows:
        chain_children = build_chain_branch(chain_rows)
        root_children.append(_topic("业务链路用例", chain_children))

    sheet_id = _new_id("sheet")
    root = _topic(root_title, root_children if root_children else None)
    root.setdefault("structureClass", "org.xmind.ui.logic.right")

    sheet = {
        "id": sheet_id,
        "class": "sheet",
        "title": "Sheet 1",
        "rootTopic": root,
    }
    return [sheet]


def write_xmind(output_path: Path, content: list[dict]) -> None:
    """Write a minimal .xmind (zip) for XMind Zen/2020+."""
    created = datetime.now(timezone.utc).isoformat()
    active_sheet_id = content[0]["id"] if content and isinstance(content, list) else None
    metadata = {
        "creator": {
            "name": "md_testcases_to_xmind.py",
            "version": "2.0",
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
    md_files: list[str]
    output: str | None
    group_by: str


def main() -> None:
    parser = argparse.ArgumentParser(description="将测试用例 Markdown 表格导出为 XMind（支持多文件合并）")
    parser.add_argument("md_files", nargs="+", help="一个或多个测试用例 Markdown 文件路径（.md）")
    parser.add_argument("-o", "--output", default=None, help="输出 .xmind 路径，默认与第一个 MD 同目录")
    parser.add_argument(
        "--group-by",
        default="模块",
        choices=["优先级", "类型", "模块"],
        help="功能点用例分组维度（默认：模块）。业务链路用例始终按链路类型分组。",
    )
    ns = parser.parse_args()
    args = Args(md_files=ns.md_files, output=ns.output, group_by=ns.group_by)

    # Parse all input files and aggregate rows
    all_functional: list[dict] = []
    all_chain: list[dict] = []

    for f in args.md_files:
        md_path = Path(f)
        if not md_path.exists():
            print(f"错误: 文件不存在 {md_path}", file=sys.stderr)
            sys.exit(1)
        if md_path.suffix.lower() != ".md":
            print(f"错误: 仅支持 .md 输入: {md_path}", file=sys.stderr)
            sys.exit(1)
        try:
            func_rows, chain_rows = load_tables_from_md(md_path)
            all_functional.extend(func_rows)
            all_chain.extend(chain_rows)
        except Exception as e:
            print(f"错误: {e}", file=sys.stderr)
            sys.exit(1)

    if not all_functional and not all_chain:
        print("错误: 所有输入文件中均未找到测试用例表格。", file=sys.stderr)
        sys.exit(1)

    # Deduplicate: same ID in functional or same 链路ID in chain
    seen_func_ids = set()
    deduped_functional = []
    for r in all_functional:
        rid = _norm(r.get("ID", ""))
        if rid and rid not in seen_func_ids:
            seen_func_ids.add(rid)
            deduped_functional.append(r)

    seen_chain_ids = set()
    deduped_chain = []
    for r in all_chain:
        rid = _norm(r.get("链路ID", ""))
        if rid and rid not in seen_chain_ids:
            seen_chain_ids.add(rid)
            deduped_chain.append(r)

    # Derive root title from first file
    first_path = Path(args.md_files[0])
    root_title = first_path.stem.replace("-功能点测试用例", "").replace("-业务链路测试用例", "").replace("-功能测试用例", "").replace("-测试用例", "") or first_path.stem

    content = build_xmind_content(deduped_functional, deduped_chain, root_title=root_title, group_by=args.group_by)

    out = Path(args.output) if args.output else first_path.parent / (root_title + "-功能测试用例.xmind")
    write_xmind(out, content)
    info_parts = []
    if deduped_functional:
        info_parts.append(f"功能点用例({len(deduped_functional)}条)")
    if deduped_chain:
        info_parts.append(f"业务链路用例({len(deduped_chain)}条)")
    if len(args.md_files) > 1:
        info_parts.append(f"来源: {len(args.md_files)}个文件")
    print(f"已导出: {out}  [{', '.join(info_parts)}]")


if __name__ == "__main__":
    main()
