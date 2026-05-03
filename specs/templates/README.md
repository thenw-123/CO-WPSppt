# Spec 模板库（L3）

Agent 可先选模板再填内容，减少结构跑偏。

| 文件 | 用途 |
|------|------|
| [outline-skeleton.json](outline-skeleton.json) | 仅封面 + 章节占位，适合「大纲阶段」 |
| [defense-starter.json](defense-starter.json) | 答辩故事线骨架（痛点→方案→技术→数据→展望） |

使用方式：`run-spec` 的 `specPath` 指向本目录下 JSON，或复制到 `specs/` 再改。
