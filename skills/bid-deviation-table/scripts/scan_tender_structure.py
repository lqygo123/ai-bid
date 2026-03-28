#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import subprocess
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
CHAPTER_RE = re.compile(r"^第[一二三四五六七八九十百零〇0-9]+章")
TOP_RE = re.compile(r"^[一二三四五六七八九十]+、")
SUB_RE = re.compile(r"^[（(][一二三四五六七八九十]+[）)]")
NUM_RE = re.compile(r"^[★▲]?[0-9]+[.、]")
CIRCLED_RE = re.compile(r"^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]")
KEYWORDS = ["商务和技术偏离表", "偏离表", "采购需求", "逐条", "一一对应", "响应文件格式"]


def normalize_text(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def convert_doc_to_docx(input_path: Path, temp_root: Path) -> Path:
    outdir = temp_root / "converted"
    outdir.mkdir(parents=True, exist_ok=True)
    profile = temp_root / "lo_profile"
    profile.mkdir(parents=True, exist_ok=True)
    cmd = [
        "soffice",
        "-env:UserInstallation=file://" + str(profile),
        "--invisible",
        "--headless",
        "--norestore",
        "--convert-to",
        "docx",
        "--outdir",
        str(outdir),
        str(input_path),
    ]
    subprocess.run(cmd, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    candidates = sorted(outdir.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise RuntimeError(f"Failed to convert .doc to .docx: {input_path}")
    return candidates[0]


def load_pdf_lines(input_path: Path, temp_root: Path) -> list[str]:
    txt_path = temp_root / f"{input_path.stem}.txt"
    cmd = ["pdftotext", "-layout", "-enc", "UTF-8", str(input_path), str(txt_path)]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    text = txt_path.read_text(encoding="utf-8", errors="ignore")
    return [normalize_text(line) for line in text.splitlines() if normalize_text(line)]


def load_docx_blocks_via_xml(input_path: Path) -> list[dict]:
    with zipfile.ZipFile(input_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        return []
    blocks = []
    para_no = 0
    table_no = 0
    for child in body:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            para_no += 1
            text = "".join(node.text for node in child.findall(".//w:t", NS) if node.text)
            text = normalize_text(text)
            if text:
                blocks.append(
                    {
                        "kind": "paragraph",
                        "locator": f"P{para_no}",
                        "text": text,
                    }
                )
            continue
        if tag == "tbl":
            table_no += 1
            rows = child.findall("./w:tr", NS)
            for r_idx, row in enumerate(rows, start=1):
                cells = row.findall("./w:tc", NS)
                for c_idx, cell in enumerate(cells, start=1):
                    text = "".join(node.text for node in cell.findall(".//w:t", NS) if node.text)
                    text = normalize_text(text)
                    if text:
                        blocks.append(
                            {
                                "kind": "table_cell",
                                "locator": f"T{table_no}R{r_idx}C{c_idx}",
                                "text": text,
                            }
                        )
    return blocks


def load_lines(input_path: Path) -> tuple[list[dict], str]:
    suffix = input_path.suffix.lower()
    with tempfile.TemporaryDirectory(prefix="scan_tender_") as tmp:
        temp_root = Path(tmp)
        source_path = input_path
        if suffix == ".doc":
            source_path = convert_doc_to_docx(input_path, temp_root)
            suffix = ".docx"
        if suffix == ".pdf":
            return (
                [{"kind": "pdf_line", "locator": f"L{i+1}", "text": row} for i, row in enumerate(load_pdf_lines(input_path, temp_root))],
                "pdf_text",
            )
        if suffix != ".docx":
            raise ValueError(f"Unsupported file type: {input_path.suffix}")
        return load_docx_blocks_via_xml(source_path), "docx_xml_blocks"


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
    if any(keyword in text for keyword in KEYWORDS):
        markers.append("keyword")
    return markers


def build_output(input_path: Path) -> dict:
    blocks, source_mode = load_lines(input_path)
    line_rows = []
    chapter_candidates = []
    keyword_hits = []
    for idx, block in enumerate(blocks, start=1):
        text = block["text"]
        markers = markers_for(text)
        line_rows.append(
            {
                "line_no": idx,
                "kind": block["kind"],
                "locator": block["locator"],
                "text": text,
                "markers": markers,
            }
        )
        if "chapter" in markers:
            chapter_candidates.append({"line_no": idx, "locator": block["locator"], "text": text})
        if "keyword" in markers:
            keyword_hits.append({"line_no": idx, "locator": block["locator"], "text": text, "markers": markers})
    return {
        "input_path": str(input_path),
        "source_mode": source_mode,
        "line_count": len(blocks),
        "chapters": chapter_candidates,
        "keyword_hits": keyword_hits,
        "lines": line_rows,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Generic tender structure scanner for deviation-table workflow")
    parser.add_argument("input_path", type=Path)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = build_output(args.input_path)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    print(f"input: {result['input_path']}")
    print(f"source_mode: {result['source_mode']}")
    print(f"line_count: {result['line_count']}")
    print("chapters:")
    for row in result["chapters"]:
        print(f"  {row['line_no']}: {row['text']}")
    print("keyword_hits:")
    for row in result["keyword_hits"][:20]:
        print(f"  {row['line_no']}: {row['text']}")


if __name__ == "__main__":
    main()
