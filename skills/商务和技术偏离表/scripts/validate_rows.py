#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path


def load_rows(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("rows.json must be a JSON array")
    return data


def looks_too_short(text: str) -> bool:
    stripped = text.strip()
    return len(stripped) <= 4


def looks_like_heading(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return True
    heading_markers = ("一、", "二、", "三、", "四、", "五、", "六、", "（一）", "（二）", "（三）")
    if stripped in heading_markers:
        return True
    if "：" not in stripped and stripped.endswith(("要求", "说明", "部分", "模块")) and len(stripped) <= 12:
        return True
    return False


def validate(rows: list[dict]) -> dict:
    findings = []
    seen_sections = set()
    required_fields = {
        "section",
        "index",
        "content",
        "requirement_text",
        "commitment_text",
        "deviation_note",
        "source_text",
        "needs_manual_review",
    }
    for i, row in enumerate(rows, start=1):
        extra_fields = sorted(set(row.keys()) - required_fields)
        missing_fields = sorted(required_fields - set(row.keys()))
        section = row.get("section", "")
        content = row.get("content", "")
        requirement = row.get("requirement_text", "")
        seen_sections.add(section)

        if "row_type" in row or "cells" in row:
            findings.append({"row": i, "type": "template_coupling", "message": "row still contains template structure fields"})
        if missing_fields:
            findings.append({"row": i, "type": "missing_fields", "message": f"missing fields: {', '.join(missing_fields)}"})
        if extra_fields:
            findings.append({"row": i, "type": "extra_fields", "message": f"unexpected fields: {', '.join(extra_fields)}"})
        if looks_too_short(requirement):
            findings.append({"row": i, "type": "short_requirement", "message": "requirement_text is very short"})
        if looks_like_heading(content):
            findings.append({"row": i, "type": "heading_like_content", "message": "content may still be a pure heading"})
        if not requirement.strip():
            findings.append({"row": i, "type": "empty_requirement", "message": "requirement_text is empty"})
        if not row.get("commitment_text", "").strip():
            findings.append({"row": i, "type": "empty_commitment", "message": "commitment_text is empty"})
        if section.strip() == content.strip() and section.strip():
            findings.append({"row": i, "type": "section_leaked_into_content", "message": "content repeats section title"})

    return {
        "row_count": len(rows),
        "sections": sorted(seen_sections),
        "findings": findings,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate semantic split quality for deviation-table rows")
    parser.add_argument("rows_path", type=Path)
    args = parser.parse_args()
    report = validate(load_rows(args.rows_path))
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
