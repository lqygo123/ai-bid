"""Shared body-walking helpers for bid-key-info-extractor.

The scanner and renderer must agree on locator semantics, otherwise highlights
land on the wrong paragraph. This module is the single source of truth.

Locator format:
    P{n}            -- the n-th <w:p> child of <w:body> (1-based, in body order)
    T{m}R{r}C{c}    -- the m-th <w:tbl> child of <w:body>, r-th row, c-th cell
"""
from __future__ import annotations

import re
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = NS["w"]


def normalize_text(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def local_tag(elem) -> str:
    tag = elem.tag
    if isinstance(tag, str) and "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return str(tag)


def gather_text(elem) -> str:
    parts = []
    for node in elem.iter(f"{{{W_NS}}}t"):
        if node.text:
            parts.append(node.text)
    return normalize_text("".join(parts))


def iter_body_blocks(body_elem):
    """Yield blocks in body order. Caller provides the <w:body> element.

    Each yield is a dict:
      {kind: "paragraph"|"table_cell",
       locator: "P5" | "T2R3C1",
       text: normalized text,
       element: the live <w:p> element to mutate,
       paragraph_index: int (1-based), only for paragraph
       table_index/row_index/cell_index: ints (1-based), only for table_cell}
    """
    para_no = 0
    table_no = 0
    for child in list(body_elem):
        tag = local_tag(child)
        if tag == "p":
            para_no += 1
            text = gather_text(child)
            yield {
                "kind": "paragraph",
                "locator": f"P{para_no}",
                "text": text,
                "element": child,
                "paragraph_index": para_no,
            }
        elif tag == "tbl":
            table_no += 1
            rows = child.findall("./w:tr", NS)
            for r_idx, row in enumerate(rows, start=1):
                cells = row.findall("./w:tc", NS)
                for c_idx, cell in enumerate(cells, start=1):
                    cell_paras = cell.findall("./w:p", NS)
                    text = normalize_text(
                        " ".join(filter(None, (gather_text(p) for p in cell_paras)))
                    )
                    yield {
                        "kind": "table_cell",
                        "locator": f"T{table_no}R{r_idx}C{c_idx}",
                        "text": text,
                        "element": cell,
                        "table_index": table_no,
                        "row_index": r_idx,
                        "cell_index": c_idx,
                        "cell_paragraphs": cell_paras,
                    }


def load_docx_body_via_xml(docx_path: Path):
    """Read <w:body> from a .docx file using zipfile + ElementTree."""
    with zipfile.ZipFile(docx_path) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find("w:body", NS)
    return body


def convert_doc_to_docx(input_path: Path, temp_root: Path) -> Path:
    """Convert .doc -> .docx via LibreOffice. Raises if soffice unavailable."""
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
    try:
        subprocess.run(cmd, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except FileNotFoundError as e:
        raise RuntimeError(
            "soffice (LibreOffice) is required to read .doc files. "
            "Install LibreOffice and ensure 'soffice' is on PATH."
        ) from e
    candidates = sorted(outdir.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise RuntimeError(f"Failed to convert .doc to .docx: {input_path}")
    return candidates[0]
