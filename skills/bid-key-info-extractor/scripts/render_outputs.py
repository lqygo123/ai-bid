#!/usr/bin/env python3
"""Render bid-key-info outputs from extract.json + source docx.

Produces:
  - <project>-招标重点.md   markdown summary
  - <stem>-标注.docx        copy of the source with yellow highlights
  - render-report.json      stats + unresolved quote list

Highlighting algorithm: anchor each evidence quote inside the paragraph or
table-cell named by its locator, splitting runs at the match boundaries
where necessary.
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from copy import deepcopy
from datetime import date
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from docx import Document  # type: ignore

from _docx_blocks import (
    W_NS,
    convert_doc_to_docx,
    iter_body_blocks,
    local_tag,
    normalize_text,
)

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

DISQUALIFY_SUBTYPES = [
    "显式否决",
    "资格审查",
    "符合性审查",
    "实质性条款",
    "散落",
]

KEY_INFO_CATEGORY_ORDER = [
    "中标人数量",
    "时间节点",
    "保证金",
    "资质要求",
    "人员社保证明",
    "其它硬性规定",
]


# -------------------------- text + index helpers --------------------------


def collect_paragraph_runs(p_elem):
    """Return ordered list of <w:r> runs inside a paragraph (incl. those
    nested in <w:hyperlink> / <w:smartTag> wrappers)."""
    return list(p_elem.iter(f"{{{W_NS}}}r"))


def run_plain_text(run_elem) -> str:
    """Return the run's searchable text. Must match the scanner's
    `_docx_blocks.gather_text` semantics: only <w:t> contents, no tab/br
    placeholders. Otherwise the agent's verbatim quote (taken from the
    scan) won't anchor here.
    """
    parts = []
    for t in run_elem.iter(f"{{{W_NS}}}t"):
        if t.text:
            parts.append(t.text)
    return "".join(parts)


def run_has_only_text(run_elem) -> bool:
    for child in run_elem:
        tag = local_tag(child)
        if tag in ("rPr", "t"):
            continue
        return False
    return True


def build_normalize_map(raw: str):
    """Return (normalized_text, raw_to_norm_index_map).
    raw_to_norm_index_map[i] = index in normalized of raw[i], or -1 if dropped.

    Normalization rules must match _docx_blocks.normalize_text:
      - \u3000 -> space
      - runs of whitespace -> single space
      - leading/trailing whitespace stripped
    """
    norm_chars: list[str] = []
    raw_to_norm: list[int] = [-1] * len(raw)
    last_was_space = True  # so leading whitespace gets dropped
    norm_to_raw_first: list[int] = []
    for i, ch in enumerate(raw):
        is_space = ch.isspace() or ch == "\u3000"
        if is_space:
            if last_was_space:
                continue
            norm_chars.append(" ")
            norm_to_raw_first.append(i)
            raw_to_norm[i] = len(norm_chars) - 1
            last_was_space = True
        else:
            norm_chars.append(ch)
            norm_to_raw_first.append(i)
            raw_to_norm[i] = len(norm_chars) - 1
            last_was_space = False
    # strip trailing space
    while norm_chars and norm_chars[-1] == " ":
        # find which raw index produced this trailing space and clear it
        last_norm = len(norm_chars) - 1
        for j in range(len(raw) - 1, -1, -1):
            if raw_to_norm[j] == last_norm:
                raw_to_norm[j] = -1
        norm_chars.pop()
        norm_to_raw_first.pop()
    return "".join(norm_chars), raw_to_norm, norm_to_raw_first


def collect_runs_with_offsets(p_elems):
    """Build a sequential run list across a list of paragraphs, joining
    paragraph boundaries with a virtual single-space separator (mirroring
    what iter_body_blocks does for cells).

    Returns (raw_text, runs_with_offsets). runs_with_offsets is
    [(run_elem, start, end), ...] in document order.
    """
    pieces = []
    runs_with_offsets = []
    cursor = 0
    for p_idx, p in enumerate(p_elems):
        if p_idx > 0:
            pieces.append(" ")
            cursor += 1
        for r in collect_paragraph_runs(p):
            rt = run_plain_text(r)
            runs_with_offsets.append((r, cursor, cursor + len(rt)))
            pieces.append(rt)
            cursor += len(rt)
    return "".join(pieces), runs_with_offsets


def find_quote_in_runs(raw: str, quote: str):
    """Locate `quote` inside the concatenated raw text.

    Returns (raw_start, raw_end) or None. Tries verbatim first, then
    normalized search with index map.
    """
    quote_norm = normalize_text(quote)
    if not quote_norm:
        return None

    idx = raw.find(quote)
    if idx != -1:
        return idx, idx + len(quote)

    raw_norm, raw_to_norm, _ = build_normalize_map(raw)
    nidx = raw_norm.find(quote_norm)
    if nidx == -1:
        return None
    nend = nidx + len(quote_norm)
    rs = None
    re_ = None
    for i, mapped in enumerate(raw_to_norm):
        if mapped == nidx and rs is None:
            rs = i
        if mapped == nend - 1:
            re_ = i + 1
    if rs is None:
        return None
    if re_ is None:
        re_ = len(raw)
    return rs, re_


# -------------------------- highlight application --------------------------


def ensure_rPr_first(run_elem):
    rPr = run_elem.find(f"{{{W_NS}}}rPr")
    if rPr is None:
        rPr = run_elem.makeelement(f"{{{W_NS}}}rPr", {})
        run_elem.insert(0, rPr)
    elif list(run_elem)[0] is not rPr:
        run_elem.remove(rPr)
        run_elem.insert(0, rPr)
    return rPr


def set_run_highlight_yellow(run_elem) -> None:
    rPr = ensure_rPr_first(run_elem)
    highlight = rPr.find(f"{{{W_NS}}}highlight")
    if highlight is None:
        highlight = rPr.makeelement(f"{{{W_NS}}}highlight", {})
        rPr.append(highlight)
    highlight.set(f"{{{W_NS}}}val", "yellow")


def clone_run_with_text(template_run, text: str, highlight: bool):
    """Create a sibling-shaped <w:r> with rPr copied from template_run and
    a single <w:t xml:space=preserve>text</w:t> child."""
    new_r = deepcopy(template_run)
    for child in list(new_r):
        if local_tag(child) != "rPr":
            new_r.remove(child)
    t = new_r.makeelement(f"{{{W_NS}}}t", {})
    t.text = text
    t.set(XML_SPACE, "preserve")
    new_r.append(t)
    if highlight:
        set_run_highlight_yellow(new_r)
    return new_r


def split_text_run_for_highlight(run_elem, run_start: int, run_end: int,
                                 hl_start: int, hl_end: int) -> bool:
    """Split a text-only run to highlight the [hl_start, hl_end) range
    (relative to the paragraph). run_start/run_end are this run's bounds.

    Returns True if the run was modified successfully.
    """
    # local offsets within the run
    local_start = max(0, hl_start - run_start)
    local_end = min(run_end - run_start, hl_end - run_start)
    if local_end <= local_start:
        return False
    full = run_plain_text(run_elem)
    before = full[:local_start]
    inside = full[local_start:local_end]
    after = full[local_end:]

    parent = run_elem.getparent()
    idx = list(parent).index(run_elem)

    new_runs = []
    if before:
        new_runs.append(clone_run_with_text(run_elem, before, highlight=False))
    if inside:
        new_runs.append(clone_run_with_text(run_elem, inside, highlight=True))
    if after:
        new_runs.append(clone_run_with_text(run_elem, after, highlight=False))

    parent.remove(run_elem)
    for i, r in enumerate(new_runs):
        parent.insert(idx + i, r)
    return True


def highlight_quote_across_runs(p_elems, quote: str) -> bool:
    """Highlight `quote` across one or more paragraphs (used for cells).

    p_elems: list of <w:p> elements in document order. Single-element lists
    are the normal paragraph case; multi-element lists are cells.
    """
    raw, runs_with_offsets = collect_runs_with_offsets(p_elems)
    found = find_quote_in_runs(raw, quote)
    if found is None:
        return False
    raw_start, raw_end = found

    for run_elem, rs, re_ in reversed(runs_with_offsets):
        if re_ <= raw_start or rs >= raw_end:
            continue
        if rs >= raw_start and re_ <= raw_end:
            set_run_highlight_yellow(run_elem)
            continue
        if run_has_only_text(run_elem):
            split_text_run_for_highlight(run_elem, rs, re_, raw_start, raw_end)
        else:
            set_run_highlight_yellow(run_elem)
    return True


def highlight_quote_in_paragraph(p_elem, quote: str) -> bool:
    return highlight_quote_across_runs([p_elem], quote)


# -------------------------- locator -> element index --------------------------


def build_locator_index(doc):
    """Map locator string -> element to highlight in.
       For paragraphs: locator -> p_elem.
       For table cells: locator -> list of p_elems inside the cell.
    """
    index = {}
    for block in iter_body_blocks(doc.element.body):
        if block["kind"] == "paragraph":
            index[block["locator"]] = ("paragraph", block["element"])
        else:
            index[block["locator"]] = ("table_cell", block["cell_paragraphs"])
    return index


def apply_evidence_to_doc(doc, evidence_items):
    """Apply highlights for a list of {locator, text_quote} entries."""
    locator_index = build_locator_index(doc)
    highlighted = 0
    unresolved = []
    for item in evidence_items:
        loc = item.get("locator", "")
        quote = item.get("text_quote", "")
        if not loc or not quote:
            unresolved.append({**item, "reason": "missing locator or text_quote"})
            continue
        target = locator_index.get(loc)
        if target is None:
            unresolved.append({**item, "reason": f"locator {loc} not found in document"})
            continue
        kind, payload = target
        if kind == "paragraph":
            ok = highlight_quote_across_runs([payload], quote)
            if ok:
                highlighted += 1
            else:
                unresolved.append({**item, "reason": f"quote not found in {loc}"})
        else:
            ok = highlight_quote_across_runs(payload, quote)
            if ok:
                highlighted += 1
            else:
                unresolved.append({**item, "reason": f"quote not found in cell {loc}"})
    return highlighted, unresolved


# -------------------------- markdown summary --------------------------


def _md_cell(text: str) -> str:
    """Escape a string for safe placement inside a markdown table cell."""
    if not text:
        return ""
    s = text.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    s = s.replace("|", "\\|")
    return s.strip()


def _render_disqualify_table(disq_evidence: list[dict]) -> list[str]:
    """Single table, sorted by subtype priority then by original index so
    rows visually group by 分类 while staying traceable to extract.json."""
    priority = {st: i for i, st in enumerate(DISQUALIFY_SUBTYPES)}
    ordered = sorted(
        enumerate(disq_evidence),
        key=lambda pair: (priority.get(pair[1].get("subtype", ""), 999), pair[0]),
    )
    lines = [
        "| # | 分类 | 要点 | 情形（原文） | 依据 |",
        "|---|------|------|------|------|",
    ]
    for row_no, (_, ev) in enumerate(ordered, 1):
        subtype = ev.get("subtype") or "未分类"
        key_point = _md_cell(ev.get("key_point", ""))
        quote = _md_cell(ev.get("text_quote", ""))
        loc = ev.get("locator", "")
        lines.append(f"| {row_no} | {subtype} | {key_point} | {quote} | `{loc}` |")
    return lines


def _render_key_info_section(idx: int, cat: str, item: dict) -> list[str]:
    lines = [f"### 2.{idx} {cat}", ""]
    summary = (item.get("summary") or "").strip()
    if summary:
        lines.append(summary)
        lines.append("")
    evidence = item.get("evidence") or []
    if evidence:
        lines.append("| 要点 | 依据 |")
        lines.append("|------|------|")
        for ev in evidence:
            quote = _md_cell(ev.get("text_quote", ""))
            loc = ev.get("locator", "")
            lines.append(f"| {quote} | `{loc}` |")
        lines.append("")
    return lines


def render_summary_markdown(extract: dict) -> str:
    meta = extract.get("project_meta") or {}
    project_name = meta.get("project_name") or "（未识别项目名）"
    project_code = meta.get("project_code") or ""
    today = date.today().isoformat()
    lines = [
        f"# 招标重点信息提要 — {project_name}",
        "",
    ]
    if project_code:
        lines.append(f"- 项目编号：{project_code}")
    lines.append(f"- 源文件：{extract.get('input_path', '')}")
    lines.append(f"- 生成日期：{today}")
    lines.append("")

    # ---- 废标项（重头戏）----
    disqualify = extract.get("废标项") or {}
    disq_summary = (disqualify.get("summary") or "").strip()
    disq_evidence = disqualify.get("evidence") or []

    lines.append(f"## 1. 废标项（共 {len(disq_evidence)} 条）")
    lines.append("")
    if disq_summary:
        lines.append(disq_summary)
        lines.append("")

    if disq_evidence:
        subtype_counts: dict[str, int] = {}
        for ev in disq_evidence:
            st = ev.get("subtype") or "未分类"
            subtype_counts[st] = subtype_counts.get(st, 0) + 1
        dist_parts = [f"{st} {subtype_counts[st]}"
                      for st in DISQUALIFY_SUBTYPES if st in subtype_counts]
        for st, cnt in subtype_counts.items():
            if st not in DISQUALIFY_SUBTYPES:
                dist_parts.append(f"{st} {cnt}")
        lines.append(f"分布：{' / '.join(dist_parts)}")
        lines.append("")
        lines.extend(_render_disqualify_table(disq_evidence))
        lines.append("")
    else:
        lines.append("_本招标文件中未发现废标触发条款（罕见，建议人工复核）。_")
        lines.append("")

    # ---- 重点信息（按标准顺序排，未识别到的 category 不出现）----
    key_info = extract.get("重点信息") or []
    by_category: dict[str, dict] = {}
    extra_categories: list[dict] = []
    for item in key_info:
        cat = item.get("category", "")
        if cat in KEY_INFO_CATEGORY_ORDER and cat not in by_category:
            by_category[cat] = item
        else:
            extra_categories.append(item)

    if by_category or extra_categories:
        lines.append("## 2. 重点信息")
        lines.append("")
        idx = 0
        for cat in KEY_INFO_CATEGORY_ORDER:
            if cat not in by_category:
                continue
            idx += 1
            lines.extend(_render_key_info_section(idx, cat, by_category[cat]))
        for item in extra_categories:
            idx += 1
            cat = item.get("category", "（未命名类别）")
            lines.extend(_render_key_info_section(idx, cat, item))

    return "\n".join(lines).rstrip() + "\n"


def collect_all_evidence(extract: dict) -> list[dict]:
    """Flatten 废标项 + 重点信息 evidence into a single list for highlighting.

    Each item gets a `_source` tag for diagnostics in the unresolved list.
    """
    items: list[dict] = []
    disqualify = extract.get("废标项") or {}
    for ev in disqualify.get("evidence") or []:
        items.append({**ev, "_source": "废标项"})
    for kf in extract.get("重点信息") or []:
        cat = kf.get("category", "")
        for ev in kf.get("evidence") or []:
            items.append({**ev, "_source": f"重点信息/{cat}"})
    return items


# -------------------------- main pipeline --------------------------


def resolve_source_for_render(source_path: Path, converted_docx: Path | None,
                              tmp_dir: Path) -> Path:
    suffix = source_path.suffix.lower()
    if suffix == ".docx":
        return source_path
    if suffix == ".doc":
        if converted_docx is not None:
            return converted_docx
        return convert_doc_to_docx(source_path, tmp_dir)
    raise ValueError(f"Unsupported source for rendering: {source_path}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Render summary.md and 标注.docx from extract.json.")
    parser.add_argument("--extract", type=Path, required=True)
    parser.add_argument("--source", type=Path, required=True, help="Original tender file (.docx or .doc)")
    parser.add_argument("--summary", type=Path, required=True)
    parser.add_argument("--highlighted", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--converted-docx", type=Path, default=None,
                        help="Pre-converted .docx if --source is .doc; reused for locator alignment")
    parser.add_argument("--strict", action="store_true",
                        help="Fail if any evidence cannot be anchored")
    args = parser.parse_args()

    extract = json.loads(args.extract.read_text(encoding="utf-8"))

    # Output dirs
    args.summary.parent.mkdir(parents=True, exist_ok=True)
    args.highlighted.parent.mkdir(parents=True, exist_ok=True)
    args.report.parent.mkdir(parents=True, exist_ok=True)

    # Summary
    md = render_summary_markdown(extract)
    args.summary.write_text(md, encoding="utf-8")

    with tempfile.TemporaryDirectory(prefix="render_bid_key_") as tmp:
        tmp_root = Path(tmp)
        source_for_render = resolve_source_for_render(args.source, args.converted_docx, tmp_root)
        doc = Document(str(source_for_render))

        all_evidence = collect_all_evidence(extract)
        highlighted, unresolved = apply_evidence_to_doc(doc, all_evidence)
        doc.save(str(args.highlighted))

    disqualify_evidence = (extract.get("废标项") or {}).get("evidence") or []
    subtype_counts: dict[str, int] = {}
    for ev in disqualify_evidence:
        st = ev.get("subtype") or "未分类"
        subtype_counts[st] = subtype_counts.get(st, 0) + 1

    key_info_categories = [
        item.get("category", "")
        for item in (extract.get("重点信息") or [])
    ]

    report = {
        "input_path": str(args.source),
        "extract_path": str(args.extract),
        "summary_path": str(args.summary),
        "highlighted_path": str(args.highlighted),
        "evidence_total": len(all_evidence),
        "highlighted": highlighted,
        "unresolved": unresolved,
        "disqualify_total": len(disqualify_evidence),
        "disqualify_by_subtype": subtype_counts,
        "key_info_categories": key_info_categories,
    }
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

    print(json.dumps(
        {k: v for k, v in report.items() if k != "unresolved"} | {"unresolved_count": len(unresolved)},
        ensure_ascii=False,
        indent=2,
    ))

    if args.strict and unresolved:
        sys.exit(2)


if __name__ == "__main__":
    main()
