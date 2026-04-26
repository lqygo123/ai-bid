# Contracts

## decision.json
Per tender run semantic decision.

```json
{
  "input_path": "demo/demo2-技术标准和要求.docx",
  "table_instruction": {
    "confirmed": true,
    "notes": [
      "Use template-driven filling with fixed commitment sentence."
    ]
  },
  "source_sections": [
    {
      "name": "商务部分",
      "selection_rule": "Rows that belong to business terms in the tender."
    },
    {
      "name": "技术部分",
      "selection_rule": "Rows from the technical requirements body."
    }
  ],
  "split_strategy": {
    "merge_continuations": true,
    "emit_heading_as_content_label": true,
    "treat_heading_only_as_group_label": true
  },
  "row_groups": [
    {
      "section": "商务部分",
      "source_ranges": [
        {
          "start_line": 100,
          "end_line": 120
        }
      ],
      "heading_handling": "heading becomes content label when needed"
    }
  ],
  "merge_rules": {
    "merge_non_structural_continuations": true,
    "keep_structural_markers": true
  },
  "fill_strategy": {
    "commitment_prefix": "我司完全响应并承诺满足：",
    "default_deviation_note": "无偏离"
  },
  "notes": [
    "This decision is only for the current tender file."
  ]
}
```

## rows.json
Stable semantic payload for downstream fill steps.

`rows.json` must be template-agnostic.

```json
[
  {
    "section": "商务部分",
    "index": "1",
    "content": "服务期",
    "requirement_text": "合同签订后1年……",
    "commitment_text": "我司完全响应并承诺满足：合同签订后1年……",
    "deviation_note": "无偏离",
    "source_text": "合同签订后1年……",
    "needs_manual_review": false
  },
  {
    "section": "技术部分",
    "index": "1",
    "content": "项目概况",
    "requirement_text": "为进一步加强……",
    "commitment_text": "我司完全响应并承诺满足：为进一步加强……",
    "deviation_note": "无偏离",
    "source_text": "项目概况：为进一步加强……",
    "needs_manual_review": false
  }
]
```

## Template Fill Assumptions
- The template exposes one table whose header row semantically contains:
  - index
  - content
  - requirement
  - commitment
  - deviation
- The fill script may infer from the template:
  - which row is the header row
  - whether a section-title row template exists
  - which row is a writable data-row template
  - whether the template expects the header row to repeat after section changes
- The fill script must not depend on hard-coded table shapes, exact column counts, or demo-specific row positions.

## Guarantees
- `decision.json` is specific to one tender run.
- `rows.json` is stable across templates for the same semantic split result.
- Temporary extraction scripts may vary across runs.
- Temporary extraction scripts should consume `decision.json` and must not hide template assumptions inside the script body.
