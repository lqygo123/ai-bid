# agent.md

## 当前范围
- 仅实现 `计划.md` 第一步：提取“响应/投标/应答文件格式”并生成模板。
- 第二步及后续暂不实现。

## 关键技能
- `$文档`：读取 `.docx/.doc/.pdf`。
- `$投标模板提取`：执行“LLM判断边界 + 定向抽取”。

## 关键工具调用
```bash
# 1) 扫描候选章节（给 agent 做语义边界判定）
python3 skills/投标模板提取/scripts/template_extract.py scan /path/to/tender.docx

# 2) 按语义判定边界做确定性抽取（生成 docx + 报告）
python3 skills/投标模板提取/scripts/template_extract.py extract \
  /path/to/tender.docx \
  /path/to/output/doc/<项目名>-投标文件模板.docx \
  --start-heading "第六章 响应文件格式" \
  --start-occurrence 1 \
  --report /path/to/output/doc/<项目名>-提取报告.md

# PDF 可编辑优先（pdf2docx，失败时回退到可编辑文本）
python3 skills/投标模板提取/scripts/template_extract.py extract \
  /path/to/tender.pdf \
  /path/to/output/doc/<项目名>-投标文件模板.docx \
  --start-heading "第四章 响应文件格式" \
  --start-occurrence 1 \
  --pdf-converter pdf2docx \
  --pdf-fallback-mode text \
  --report /path/to/output/doc/<项目名>-提取报告.md
```

## 边界判定要点
- 起点必须是正文中的格式章节，不能用目录里的同名项。
- 同名章节多次出现时，用 `--start-occurrence` 指向正文那次。
- 自动截断不稳时，补 `--end-heading`（必要时 `--match-mode contains`）。
- 禁止关键词规则自动选章，边界必须由 agent 语义判断给出。
