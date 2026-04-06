#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from copy import deepcopy
from pathlib import Path

from docx import Document


INDEX_KEYS = ("序号", "序", "编号", "index")
CONTENT_KEYS = ("内容", "项目", "content")
REQUIREMENT_KEYS = ("要求", "标准", "招标文件要求", "响应要求", "requirement")
COMMITMENT_KEYS = ("承诺", "投标情况", "响应内容", "响应承诺", "commitment")
DEVIATION_KEYS = ("偏离", "偏差", "deviation")


def normalize_text(text: str) -> str:
    return " ".join(text.split())


def load_rows(rows_path: Path) -> list[dict]:
    data = json.loads(rows_path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("rows.json must be a JSON array")
    return data


def row_texts(row) -> list[str]:
    return [normalize_text(cell.text) for cell in row.cells]


def has_any(text: str, keywords: tuple[str, ...]) -> bool:
    return any(keyword in text for keyword in keywords)


def classify_header_row(texts: list[str]) -> dict | None:
    roles: dict[str, int] = {}
    for idx, text in enumerate(texts):
        if not text:
            continue
        if "index" not in roles and has_any(text, INDEX_KEYS):
            roles["index"] = idx
            continue
        if "content" not in roles and has_any(text, CONTENT_KEYS):
            roles["content"] = idx
            continue
        if "requirement" not in roles and has_any(text, REQUIREMENT_KEYS):
            roles["requirement"] = idx
            continue
        if "commitment" not in roles and has_any(text, COMMITMENT_KEYS):
            roles["commitment"] = idx
            continue
        if "deviation" not in roles and has_any(text, DEVIATION_KEYS):
            roles["deviation"] = idx
            continue
    required = {"requirement", "commitment", "deviation"}
    if required.issubset(roles):
        return roles
    return None


def is_section_like_row(texts: list[str]) -> bool:
    non_empty = [text for text in texts if text]
    if not non_empty:
        return False
    return len(set(non_empty)) == 1


def find_target_table(doc: Document):
    best = None
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            roles = classify_header_row(row_texts(row))
            if roles is None:
                continue
            best = (table, row_idx, roles)
            return best
    return None


def remove_row(table, row) -> None:
    table._tbl.remove(row._tr)


def clone_row(table, row):
    new_tr = deepcopy(row._tr)
    table._tbl.append(new_tr)
    return table.rows[-1]


def clone_template_row(table, template_tr):
    new_tr = deepcopy(template_tr)
    table._tbl.append(new_tr)
    return table.rows[-1]


def set_cell_text(cell, text: str) -> None:
    cell.text = text


def infer_templates(table, header_idx: int):
    header_template = table.rows[header_idx]
    section_template = None
    section_idx = None
    for idx in range(header_idx - 1, -1, -1):
        texts = row_texts(table.rows[idx])
        if is_section_like_row(texts):
            section_template = table.rows[idx]
            section_idx = idx
            break
        if any(texts):
            break

    data_template = None
    data_idx = None
    for idx in range(header_idx + 1, len(table.rows)):
        texts = row_texts(table.rows[idx])
        if not any(texts):
            continue
        if classify_header_row(texts) is not None:
            continue
        if is_section_like_row(texts):
            continue
        data_template = table.rows[idx]
        data_idx = idx
        break

    if data_template is None:
        raise ValueError("Could not infer a writable data template row")
    return {
        "section_template_tr": deepcopy(section_template._tr) if section_template is not None else None,
        "section_idx": section_idx,
        "header_template_tr": deepcopy(header_template._tr),
        "header_values": row_texts(header_template),
        "header_idx": header_idx,
        "data_template_tr": deepcopy(data_template._tr),
        "data_idx": data_idx,
        "repeat_header_on_section": section_template is not None,
    }


def truncate_after(table, keep_rows: int) -> None:
    while len(table.rows) > keep_rows:
        remove_row(table, table.rows[-1])


def fill_from_rows(table, header_idx: int, roles: dict[str, int], rows: list[dict]) -> None:
    templates = infer_templates(table, header_idx)
    truncate_after(table, header_idx + 1)

    current_section = None
    for row_data in rows:
        section = row_data.get("section", "")
        if section != current_section:
            current_section = section
            if templates["section_template_tr"] is not None:
                row = clone_template_row(table, templates["section_template_tr"])
                for idx in range(len(row.cells)):
                    set_cell_text(row.cells[idx], section)
            if templates["repeat_header_on_section"]:
                row = clone_template_row(table, templates["header_template_tr"])
                for idx, value in enumerate(templates["header_values"][: len(row.cells)]):
                    set_cell_text(row.cells[idx], value)

        row = clone_template_row(table, templates["data_template_tr"])
        for cell in row.cells:
            set_cell_text(cell, "")
        if "index" in roles:
            set_cell_text(row.cells[roles["index"]], str(row_data.get("index", "")))
        if "content" in roles:
            set_cell_text(row.cells[roles["content"]], row_data.get("content", ""))
            requirement_value = row_data.get("requirement_text", "")
            commitment_value = row_data.get("commitment_text", "")
        else:
            content = row_data.get("content", "").strip()
            requirement_text = row_data.get("requirement_text", "").strip()
            commitment_text = row_data.get("commitment_text", "").strip()
            requirement_value = f"{content}：{requirement_text}" if content else requirement_text
            commitment_value = f"{content}：{commitment_text}" if content else commitment_text
        set_cell_text(row.cells[roles["requirement"]], requirement_value)
        set_cell_text(row.cells[roles["commitment"]], commitment_value)
        set_cell_text(row.cells[roles["deviation"]], row_data.get("deviation_note", ""))


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill deviation table using template-driven semantic column mapping")
    parser.add_argument("template_path", type=Path)
    parser.add_argument("rows_path", type=Path)
    parser.add_argument("output_path", type=Path)
    args = parser.parse_args()

    rows = load_rows(args.rows_path)
    doc = Document(str(args.template_path))
    target = find_target_table(doc)
    if target is None:
        raise ValueError("Could not find a deviation table with semantic header columns")

    table, header_idx, roles = target
    fill_from_rows(table, header_idx, roles, rows)
    args.output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.output_path))
    print(
        json.dumps(
            {
                "template_path": str(args.template_path),
                "rows_path": str(args.rows_path),
                "output_path": str(args.output_path),
                "filled_rows": len(rows),
                "header_row_index": header_idx,
                "column_roles": roles,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
