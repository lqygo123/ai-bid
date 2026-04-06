#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from copy import deepcopy
from pathlib import Path

from docx import Document


INDEX_KEYS = ("\u5e8f\u53f7", "\u7f16\u53f7", "index")
CONTENT_KEYS = ("\u5185\u5bb9", "\u9879\u76ee", "content")
REQUIREMENT_KEYS = (
    "\u8981\u6c42",
    "\u6807\u51c6",
    "\u62db\u6807\u6587\u4ef6\u8981\u6c42",
    "\u62db\u6807\u6587\u4ef6\u6280\u672f\u8981\u6c42",
    "\u62db\u6807\u6587\u4ef6\u5546\u52a1\u8981\u6c42",
    "\u62db\u6807\u670d\u52a1\u9700\u6c42",
    "requirement",
)
COMMITMENT_KEYS = (
    "\u627f\u8bfa",
    "\u6295\u6807\u60c5\u51b5",
    "\u54cd\u5e94\u5185\u5bb9",
    "\u54cd\u5e94\u627f\u8bfa",
    "\u6295\u6807\u6587\u4ef6\u6280\u672f\u54cd\u5e94",
    "\u6295\u6807\u6587\u4ef6\u5546\u52a1\u54cd\u5e94",
    "\u6295\u6807\u670d\u52a1\u54cd\u5e94",
    "commitment",
)
DEVIATION_KEYS = ("\u504f\u79bb", "\u504f\u5dee", "deviation")


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


def remove_row(table, row) -> None:
    table._tbl.remove(row._tr)


def clone_template_row(table, template_tr):
    new_tr = deepcopy(template_tr)
    table._tbl.append(new_tr)
    return table.rows[-1]


def set_cell_text(cell, text: str) -> None:
    cell.text = text


def infer_templates(table, header_idx: int):
    header_template = table.rows[header_idx]
    data_template = None
    for idx in range(header_idx + 1, len(table.rows)):
        texts = row_texts(table.rows[idx])
        if not any(texts):
            continue
        if classify_header_row(texts) is not None:
            continue
        data_template = table.rows[idx]
        break
    if data_template is None:
        raise ValueError("Could not infer a writable data template row")
    return {
        "header_template_tr": deepcopy(header_template._tr),
        "data_template_tr": deepcopy(data_template._tr),
    }


def truncate_after(table, keep_rows: int) -> None:
    while len(table.rows) > keep_rows:
        remove_row(table, table.rows[-1])


def split_rows_by_section(rows: list[dict]) -> dict[str, list[dict]]:
    grouped: dict[str, list[dict]] = {}
    for row in rows:
        section = normalize_text(row.get("section", "")) or "__default__"
        grouped.setdefault(section, []).append(row)
    return grouped


def format_requirement_and_commitment(row_data: dict, has_content_column: bool) -> tuple[str, str]:
    if has_content_column:
        return row_data.get("requirement_text", ""), row_data.get("commitment_text", "")

    content = row_data.get("content", "").strip()
    requirement_text = row_data.get("requirement_text", "").strip()
    commitment_text = row_data.get("commitment_text", "").strip()
    if content:
        return f"{content}：{requirement_text}", f"{content}：{commitment_text}"
    return requirement_text, commitment_text


def fill_table(table, header_idx: int, roles: dict[str, int], rows: list[dict]) -> int:
    templates = infer_templates(table, header_idx)
    truncate_after(table, header_idx + 1)

    filled = 0
    has_content_column = "content" in roles
    for row_data in rows:
        row = clone_template_row(table, templates["data_template_tr"])
        for cell in row.cells:
            set_cell_text(cell, "")
        if "index" in roles:
            set_cell_text(row.cells[roles["index"]], str(row_data.get("index", "")))
        if has_content_column:
            set_cell_text(row.cells[roles["content"]], row_data.get("content", ""))
        requirement_value, commitment_value = format_requirement_and_commitment(
            row_data, has_content_column
        )
        set_cell_text(row.cells[roles["requirement"]], requirement_value)
        set_cell_text(row.cells[roles["commitment"]], commitment_value)
        set_cell_text(row.cells[roles["deviation"]], row_data.get("deviation_note", ""))
        filled += 1
    return filled


def find_target_tables(doc: Document) -> list[tuple[object, int, dict]]:
    matches = []
    for table in doc.tables:
        for row_idx, row in enumerate(table.rows):
            roles = classify_header_row(row_texts(row))
            if roles is None:
                continue
            matches.append((table, row_idx, roles))
            break
    return matches


def resolve_table_sections(section_names: list[str], table_count: int) -> list[str]:
    if table_count == 1:
        if len(section_names) == 1:
            return [section_names[0]]
        return ["__all__"]
    if len(section_names) < table_count:
        raise ValueError(
            f"Not enough sections in rows.json for template tables: sections={section_names}, tables={table_count}"
        )
    return section_names[:table_count]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fill deviation tables using template-driven semantic column mapping"
    )
    parser.add_argument("template_path", type=Path)
    parser.add_argument("rows_path", type=Path)
    parser.add_argument("output_path", type=Path)
    args = parser.parse_args()

    rows = load_rows(args.rows_path)
    doc = Document(str(args.template_path))
    targets = find_target_tables(doc)
    if not targets:
        raise ValueError("Could not find a deviation table with semantic header columns")

    grouped_rows = split_rows_by_section(rows)
    section_names = [name for name in grouped_rows.keys() if name != "__default__"]
    target_sections = resolve_table_sections(section_names, len(targets))

    summaries = []
    total_filled = 0
    for idx, (table, header_idx, roles) in enumerate(targets, start=1):
        section_name = target_sections[idx - 1]
        section_rows = rows if section_name == "__all__" else grouped_rows[section_name]
        filled = fill_table(table, header_idx, roles, section_rows)
        total_filled += filled
        summaries.append(
            {
                "table_index": idx,
                "section": section_name,
                "filled_rows": filled,
                "header_row_index": header_idx,
                "column_roles": roles,
            }
        )

    args.output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(args.output_path))
    print(
        json.dumps(
            {
                "template_path": str(args.template_path),
                "rows_path": str(args.rows_path),
                "output_path": str(args.output_path),
                "filled_rows": total_filled,
                "tables": summaries,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
