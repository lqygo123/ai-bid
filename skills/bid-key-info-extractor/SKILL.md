---
name: bid-key-info-extractor
description: 招标重点信息提取 workflow skill — 详尽抽取所有可能导致废标的条款（重头戏），并轻量提炼几类常关注的重点信息。输出黄色高亮的 .docx 供人工复核 + Markdown 概要。
---

# Bid Key-Info Extractor

两个真实需求：

1. **废标项**——所有可能导致投标被否决/作废/视为无效的条款，**全文兜底**找出来标黄。这是这个 skill 的重头戏。
2. **重点信息**——你在每份招标文件里常关注的几个点（时间节点、保证金、资质等），识别到了就引一条，没有就跳过。

## Hard rule

提取必须基于语义理解，**不能**用关键词匹配/正则当筛子。Scanner 给的 `keyword_hits` 仅作注意力提示。

## Workflow

### 1) Scan
```bash
python3 skills/bid-key-info-extractor/scripts/scan_tender_structure.py \
  /path/to/tender.docx --json > tmp/docs/<run>/scan.json
```

`scan.json` 列出每段 / 每单元格的 `{locator, text, keyword_hits}`，locator 形如 `P{n}` 或 `T{m}R{r}C{c}`。

### 2) Agent 语义抽取

读 `scan.json`（必要时也读原文），按 [references/extraction-guide.md](references/extraction-guide.md) 指引产出 `extract.json`。

**Schema**：

```json
{
  "input_path": "招标文件示例/demo.docx",
  "project_meta": {"project_name": "...", "project_code": "..."},
  "废标项": {
    "summary": "本项目废标触发情形包含 ...",
    "evidence": [
      {"locator": "P142", "text_quote": "投标人有下列情形之一的，其投标将被否决：", "subtype": "显式否决", "key_point": "总则：列出否决情形的引导句"},
      {"locator": "T3R5C2", "text_quote": "营业执照不在有效期内的", "subtype": "资格审查", "key_point": "营业执照失效"}
    ]
  },
  "重点信息": [
    {
      "category": "时间节点",
      "summary": "报名截止：2026-05-10 17:00；开标时间：2026-05-15 09:30",
      "evidence": [
        {"locator": "T1R8C2", "text_quote": "报名截止时间：2026 年 5 月 10 日 17 时"}
      ]
    },
    {
      "category": "保证金",
      "summary": "投标保证金 5 万元，电子保函或银行转账",
      "evidence": [{"locator": "T1R12C2", "text_quote": "投标保证金人民币伍万元"}]
    }
  ]
}
```

**契约要点**：

- `废标项` 是顶层固定字段，**必须存在**（即使罕见地全文无废标条款，也写空 evidence + 备注 summary）。
- `重点信息` 是数组。**只放找到的类别，没找到的就不出现**——不要写 `present=false` 占位。
- `category` 取自这 6 个建议值：`中标人数量` / `时间节点` / `保证金` / `资质要求` / `人员社保证明` / `其它硬性规定`。同一 category 不要重复出现，整组 evidence 合并到一个对象里。
- 每条 evidence 的 `text_quote` 必须是 locator 文本的 **verbatim 子串**（按 `normalize_text` 比较：`\u3000`→空格、连续空白合并、首尾去空）。单条 ≤ 200 字符，不可跨段落。
- `废标项.evidence` 的每项必须带 `subtype`（取值见 §1 五类）和 `key_point`（≤ 30 字的浓缩要点，渲染时进表格的"要点"列）。`key_point` 是 agent 复述，不需要 verbatim。
- `summary` 是 agent 自由复述，要的是可读性；`text_quote` 是锚点，要的是可锚定；`key_point` 是表格抓手，要的是一眼能看懂这一行讲什么。

### 3) Render

```bash
python3 skills/bid-key-info-extractor/scripts/render_outputs.py \
  --extract tmp/docs/<run>/extract.json \
  --source /path/to/tender.docx \
  --summary output/doc/<项目名>-招标重点.md \
  --highlighted output/doc/<项目名>-标注.docx \
  --report tmp/docs/<run>/render-report.json
```

`.doc` 输入：传 `--converted-docx tmp/docs/<run>/converted.docx`（复用 scanner 转出的同一份）。
`--strict` 在任何 evidence 锚不上时直接失败。

`render-report.json.unresolved` 非空 = quote 没锚上，多半是跨段引用或被改写过——改 `extract.json` 重渲染。

## 抽取重点

**废标项是这个 skill 的重头**，单独看 [extraction-guide.md §1](references/extraction-guide.md)，覆盖五类：显式否决条款 / 资格审查不通过 / 符合性审查不通过 / 实质性条款（★/▲）/ 散落的"否则无效"句。典型采购文件 15-40 条，**少于 10 条**几乎一定是漏标。

**其他 6 类轻量处理**：识别到就引 1-3 条最有代表性的 evidence，写一行 summary 说清要点；没识别到就不出现。不强求覆盖、不强求穷举。

## Failure rules

- Quote 锚不上：改 `extract.json`，让每条 quote 是 locator 文本的 verbatim 子串。
- `.doc` 转换失败：先装 LibreOffice（`soffice` on PATH）。
- 单条 quote > 200 字符：拆成多条。

## Working conventions

- 中间产物：`tmp/docs/<run-name>/scan.json`、`extract.json`、`render-report.json`。
- 最终产物：`output/doc/<项目名>-招标重点.md`、`output/doc/<项目名>-标注.docx`。
- `<项目名>` 取自 `extract.project_meta.project_name`，缺则用源文件名 stem。

## Out of scope

- PDF 输入（renderer 不支持高亮 PDF）。
- 多色分类高亮（仅黄色）。
- 多轮/分块抽取（单次 LLM 调用足以处理常见招标文件）。

## Script reference

| Script | Purpose |
|---|---|
| `scripts/scan_tender_structure.py` | Body-order scan with stable locators. |
| `scripts/render_outputs.py` | summary.md + 标注.docx + report. |
| `scripts/_docx_blocks.py` | 共享 XML 遍历。Scanner 与 renderer 必须用同一份遍历逻辑，否则 locator 对不上。 |

依赖：`python-docx`；`soffice`（仅 `.doc` 输入需要）。
