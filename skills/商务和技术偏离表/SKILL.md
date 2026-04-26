---
name: 商务和技术偏离表
description: Workflow skill for 商务和技术偏离表 using semantic splitting plus template-driven fixed-commitment filling. Use when you need to read one tender file, decide row splitting semantically for that file, generate rows.json, and write those rows back into a deviation-table template without coupling to a specific table shape.
---

# 商务和技术偏离表

This skill is a workflow skill, not a fixed parser and not a fixed table-format adapter.

## Intent
Use it when a tender file contains a deviation table and the target fill style is simple:
- source requirements are copied into the table
- commitments follow a fixed sentence pattern
- deviation notes are usually `无偏离`

The durable part is the workflow:
1. Read the tender file.
2. Semantically decide how this file should be split into rows.
3. Write a temporary extractor script for this run.
4. Produce `rows.json`.
5. Fill the target template from `rows.json` using template structure inference.

Do not keep adding one-off heuristics into the reusable skill scripts.

## Stable Filling Style
This skill only keeps the `demo2` style of filling:
- template-driven
- fixed commitment sentence
- semantic row splitting still required

Recommended field semantics:
- requirement text: copy the tender requirement with minimal rewriting
- commitment text: `我司完全响应并承诺满足：` + requirement text
- deviation note: default `无偏离`
- rows.json: semantic-only payload, independent from the template's visual row layout

What still needs semantic judgment per file:
- which chapter or section feeds the table
- how coarse or fine each row should be
- which headings should become `内容` labels
- which requirement paragraphs should be merged into one row
- where section breaks such as `商务部分` / `技术部分` should occur

## Durable vs Temporary
### Durable assets in this skill
- generic document scanning
- decision contract
- rows contract
- template-driven fill contract
- split-result review checklist

### Temporary assets per tender run
- `decision.json`
- `extract_rows.py`
- `rows.json`
- optional debug artifacts

Write those run-specific files under `tmp/docs/<run-name>/`.

## Workflow
### 1. Scan the file structure
Run the generic scanner first.

```bash
python skills/商务和技术偏离表/scripts/scan_tender_structure.py \
  /path/to/tender.docx \
  --json > /tmp/scan.json
```

Use the scan to inspect:
- chapter boundaries
- numbering style
- candidate lines near the deviation-table instructions
- candidate source sections
- paragraph and table-cell text together, in document order when possible

### 2. Make a semantic decision for this file
Before writing any extractor, produce `decision.json` for this tender.

The decision must lock:
- source section
- row split granularity
- heading vs row treatment
- continuation merge rules
- section grouping rules
- fixed commitment prefix

Use the contract in [references/contracts.md](references/contracts.md).

### 3. Write a temporary extractor script
Write a one-off script under the run directory, for example:

```bash
tmp/docs/my-run/extract_rows.py
```

That script may assume the current file's numbering and continuation style.
It should be short and explicit. It should not modify reusable skill code.
It should consume the decisions already written in `decision.json`, rather than re-hiding those judgments in code.

### 4. Run the temporary script
The temporary script should emit `rows.json` using the stable contract.

### 4.5 Review the split result
Review `rows.json` before filling.

Quick check:

```bash
python skills/商务和技术偏离表/scripts/validate_rows.py \
  /path/to/rows.json
```

Minimum checks:
- every `content` is independently understandable
- every `requirement_text` still makes sense outside its original paragraph context
- no pure heading row was emitted as data
- no important continuation paragraph was dropped
- business and technical sections cover the intended source requirements

### 5. Fill the template
Run the generic fill script:

```bash
python skills/商务和技术偏离表/scripts/fill_deviation_table.py \
  /path/to/template.docx \
  /path/to/rows.json \
  /path/to/output.docx
```

The fill script should infer the template structure from the template itself.
It must not branch on hard-coded table counts, fixed row positions, or demo-specific table layouts.

## Contracts
Read [references/contracts.md](references/contracts.md).

Stable output for downstream use:
- `decision.json`
- `rows.json`

The contracts are durable.
The extraction logic is not.

## Failure Rules
- If the tender does not clearly say where the table content comes from, stop and mark it for manual review.
- If multiple source sections are plausible, keep that ambiguity in `decision.json` instead of forcing a parser rule into the skill.
- If the template does not expose a recognizable header row for requirement/commitment/deviation semantics, stop and inspect the template manually instead of hard-coding another format branch.
- If scan output misses text because it only came from tables, fix the scan step first rather than compensating in the extractor.

## Working Conventions
- Reusable scripts live in `skills/商务和技术偏离表/scripts/`.
- Run artifacts live in `tmp/docs/<run-name>/`.
- Keep reusable scripts generic and inspection-oriented.
- Keep run-specific scripts disposable.
