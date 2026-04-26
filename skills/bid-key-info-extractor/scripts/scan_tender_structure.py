#!/usr/bin/env python3
"""Tender structure scanner for bid-key-info-extractor.

Emits a JSON map of every paragraph/table-cell in body order with stable
locators. Keyword markers are *attention hints only* — never used to decide
extraction. The agent reads this output and makes the semantic call.
"""
from __future__ import annotations

import argparse
import io
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
elif isinstance(sys.stdout, io.TextIOWrapper):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from _docx_blocks import (
    NS,
    convert_doc_to_docx,
    iter_body_blocks,
    load_docx_body_via_xml,
    normalize_text,
)

CHAPTER_RE = re.compile(r"^第[一二三四五六七八九十百零〇0-9]+章")
TOP_RE = re.compile(r"^[一二三四五六七八九十]+、")
SUB_RE = re.compile(r"^[（(][一二三四五六七八九十]+[）)]")
NUM_RE = re.compile(r"^[★▲]?[0-9]+[.、]")
CIRCLED_RE = re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")

KEYWORDS = [
    "废标", "无效投标", "否决", "拒绝投标", "不予受理", "作废",
    "投标保证金", "履约保证金", "保证金", "押金",
    "开标时间", "开标地点", "投标截止", "递交截止", "提交截止",
    "报名截止", "报名时间", "投标截止时间",
    "资质", "资质证书", "等级证书", "执业资格", "营业执照",
    "社保", "社会保险", "公积金", "纳税", "税务证明", "完税证明",
    "联合体", "分包", "转包",
    "入围", "入选", "中标人数量", "中标候选人", "几家",
    "项目编号", "项目名称",
]


def load_pdf_lines(input_path: Path, temp_root: Path) -> list[str]:
    txt_path = temp_root / f"{input_path.stem}.txt"
    cmd = ["pdftotext", "-layout", "-enc", "UTF-8", str(input_path), str(txt_path)]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    text = txt_path.read_text(encoding="utf-8", errors="ignore")
    return [normalize_text(line) for line in text.splitlines() if normalize_text(line)]


def load_blocks(input_path: Path) -> tuple[list[dict], str]:
    suffix = input_path.suffix.lower()
    temp_root = Path(tempfile.mkdtemp(prefix="scan_bid_key_"))
    source_path = input_path
    if suffix == ".doc":
        source_path = convert_doc_to_docx(input_path, temp_root)
        suffix = ".docx"
    if suffix == ".pdf":
        return (
            [
                {"kind": "pdf_line", "locator": f"L{i+1}", "text": row}
                for i, row in enumerate(load_pdf_lines(input_path, temp_root))
            ],
            "pdf_text",
        )
    if suffix != ".docx":
        raise ValueError(f"Unsupported file type: {input_path.suffix}")
    body = load_docx_body_via_xml(source_path)
    if body is None:
        return [], "docx_xml_blocks"
    out = []
    for block in iter_body_blocks(body):
        if not block["text"]:
            continue
        out.append(
            {
                "kind": block["kind"],
                "locator": block["locator"],
                "text": block["text"],
            }
        )
    return out, "docx_xml_blocks"


def markers_for(text: str) -> list[str]:
    markers = []
    if CHAPTER_RE.match(text):
        markers.append("chapter")
    if TOP_RE.match(text):
        markers.append("top")
    if SUB_RE.match(text):
        markers.append("sub")
    if NUM_RE.match(text):
        markers.append("num")
    if CIRCLED_RE.match(text):
        markers.append("circled")
    hit_keywords = [kw for kw in KEYWORDS if kw in text]
    if hit_keywords:
        markers.append("keyword")
    return markers, hit_keywords


def build_output(input_path: Path) -> dict:
    blocks, source_mode = load_blocks(input_path)
    line_rows = []
    chapter_candidates = []
    keyword_hits = []
    for idx, block in enumerate(blocks, start=1):
        text = block["text"]
        markers, hit_keywords = markers_for(text)
        row = {
            "line_no": idx,
            "kind": block["kind"],
            "locator": block["locator"],
            "text": text,
            "markers": markers,
        }
        if hit_keywords:
            row["keywords"] = hit_keywords
        line_rows.append(row)
        if "chapter" in markers:
            chapter_candidates.append({"line_no": idx, "locator": block["locator"], "text": text})
        if "keyword" in markers:
            keyword_hits.append(
                {
                    "line_no": idx,
                    "locator": block["locator"],
                    "text": text,
                    "keywords": hit_keywords,
                }
            )
    return {
        "input_path": str(input_path),
        "source_mode": source_mode,
        "line_count": len(blocks),
        "chapters": chapter_candidates,
        "keyword_hits": keyword_hits,
        "lines": line_rows,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Tender structure scanner for bid key-info extraction. "
        "Output is for the agent's semantic reading; markers are hints only."
    )
    parser.add_argument("input_path", type=Path)
    parser.add_argument("--json", action="store_true", help="Emit full JSON.")
    args = parser.parse_args()
    result = build_output(args.input_path)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    print(f"input: {result['input_path']}")
    print(f"source_mode: {result['source_mode']}")
    print(f"line_count: {result['line_count']}")
    print(f"chapter_count: {len(result['chapters'])}")
    print(f"keyword_hit_count: {len(result['keyword_hits'])}")
    print("chapters:")
    for row in result["chapters"]:
        print(f"  {row['line_no']:>5} {row['locator']:<10} {row['text'][:80]}")
    print("first 30 keyword hits:")
    for row in result["keyword_hits"][:30]:
        kws = ",".join(row["keywords"])
        print(f"  {row['line_no']:>5} {row['locator']:<10} [{kws}] {row['text'][:80]}")


if __name__ == "__main__":
    main()
