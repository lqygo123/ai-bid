---
name: bid-template-extractor
description: Extract bid/response document templates from tender files using LLM semantic boundary judgment plus deterministic section copying. Use when processing one .docx/.doc/.pdf tender document and you need reusable template outputs for downstream auto-fill.
---

# Bid Template Extractor

Use this skill as:
1. Agent does semantic boundary decision.
2. Script does deterministic extraction.

## Single-File Workflow

### 1) Scan markers first
Run scan and review candidate lines before extraction.

```bash
python3 scripts/template_extract.py scan /path/to/tender.docx
python3 scripts/template_extract.py scan /path/to/tender.doc
python3 scripts/template_extract.py scan /path/to/tender.pdf
```

Focus on:
- Real section start in body (not TOC lines).
- Correct occurrence index when same heading appears multiple times.

### 2) Agent semantic decision (required)
Boundary choice must come from agent semantic understanding.

Hard rule:
- Do not use keyword-only auto matching to pick boundaries.

Typical decision output:
- `start_heading`
- `start_occurrence`
- optional `end_heading` and `end_occurrence`
- `match_mode` (`exact` first, `contains` only when needed)

### 3) Deterministic extract

```bash
python3 scripts/template_extract.py extract \
  /path/to/tender.docx \
  /path/to/output-template.docx \
  --start-heading "第六章 响应文件格式" \
  --start-occurrence 1 \
  --report /path/to/extract-report.md
```

For editable PDF extraction:

```bash
python3 scripts/template_extract.py extract \
  /path/to/tender.pdf \
  /path/to/output-template.docx \
  --start-heading "第四章 响应文件格式" \
  --start-occurrence 1 \
  --pdf-converter pdf2docx \
  --pdf-fallback-mode text \
  --report /path/to/extract-report.md
```

When auto stop is ambiguous:

```bash
python3 scripts/template_extract.py extract \
  /path/to/tender.docx \
  /path/to/output-template.docx \
  --start-heading "第五章 谈判应答文件格式" \
  --start-occurrence 2 \
  --end-heading "第六章" \
  --match-mode contains \
  --report /path/to/extract-report.md
```

### 4) PDF handling
- Default converter is `pdf2docx`, so output stays editable whenever conversion succeeds.
- Optional: `--pdf-converter soffice` to switch converter order.
- If conversion fails, default fallback is `text` (still editable, lower layout fidelity).
- Optional: `--pdf-fallback-mode image_pages` for visual fidelity (not editable).
- Check report field `source_mode`:
  - `pdf_pdf2docx` or `pdf_soffice_docx`: editable structured mode.
  - `pdf_text_fallback`: editable text mode.
  - `pdf_page_images`: high-fidelity visual mode (not editable).

### 5) Quick verification
- Inspect first 10-20 non-empty paragraphs in output.
- Confirm output starts with expected format section heading.
- Confirm key parts exist (cover/catalog/response letter/报价表/资格材料).
- Check report stats: `source_mode`, `copied_paragraphs`, `copied_tables`.

## Script Reference

### scripts/template_extract.py
- `scan`: prints chapter/keyword marker rows for semantic boundary judgment.
- `extract`: copies section into standalone `.docx`.

Supported inputs:
- `.docx` direct read.
- `.doc` via `soffice`.
- `.pdf` via `pdf2docx` first (default), then `soffice`, then fallback mode (`text` by default).

Dependencies:
- `python-docx`
- `pdf2docx` (preferred PDF editable conversion path)
- `soffice` (for `.doc`, and secondary PDF conversion path)
- `pdftotext` (for `.pdf` fallback path)

## Working Conventions
- Keep intermediate files under `tmp/docs/`.
- Write final templates under `output/doc/`.
- Keep filename stable: `<项目名>-投标文件模板.docx`.
