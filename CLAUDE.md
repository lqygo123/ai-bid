# ai-bid — AI 招投标工作流

把招投标流程中的确定性操作沉淀为可复用的 skill，由 Claude Code 驱动执行。

## 工程级 Skill

本项目注册了 4 个 skill，每个 skill 的完整工作流文档在对应 `SKILL.md` 中：

| Skill 名 | 类型 | 路径 | 用途 |
|---|---|---|---|
| `bid-key-info-extractor` | 信息提取 | `skills/bid-key-info-extractor/SKILL.md` | 从招标文件抽取废标项（重头）+ 6 类重点信息，输出标黄 docx + markdown 摘要 |
| `bid-deviation-table` | 标书撰写 | `skills/商务和技术偏离表/SKILL.md` | 语义拆分 + 模板驱动的固定承诺语句偏离表填写 |
| `bid-template-extractor` | 标书撰写 | `skills/投标模板提取/SKILL.md` | LLM 语义边界判断 + 确定性章节复制，从招标文件提取投标模板 |
| `docx` | 基础工具 | `skills/文档/SKILL.md` | DOCX 读写/编辑/渲染（python-docx + LibreOffice 转 PDF 审阅） |

## 工作约定

- **临时文件**：`tmp/docs/<run-name>/`，用完可删
- **最终产物**：`output/doc/`
- **脚本**：确定性操作放 `skills/<name>/scripts/`，不做语义判断
- **Agent 职责**：语义判断（哪些是废标项、如何拆分偏离表行、选择模板边界）
- **参考文档**：`skills/<name>/references/` 下存放 Agent 推理指引

## 执行流程

每个 skill 遵循：**Scan（脚本扫结构）→ Agent 语义决策 → 确定性 Render/Fill（脚本执行）**

Agent 不做纯文本匹配，脚本不做语义判断。
