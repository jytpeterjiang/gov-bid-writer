# 政府采购投标文件编写专家（gov-bid-writer）

政府采购投标文件编写技能（AI Skill），覆盖从招标文件解读、评分策略分析、投标方案撰写到最终 docx 文档生成的全流程。

也可通过 [skillhub](https://skillhub.cn/skills/gov-bid-writer) 获取本技能。

## 环境要求

- Python 依赖：`python-docx`、`pdfplumber`
- Conda 环境名：`bid-skills`

## 工作流程

1. **招标文件解读** -- 提取采购需求、资格要求、评分规则等核心要素
2. **评分策略分析** -- 识别高分值项，制定得分策略
3. **内容撰写** -- 按标书标准结构逐条响应招标要求，生成各章节内容
4. **文档生成** -- 输出格式规范的 `.docx` 投标文件

## 关键写作原则

- **逐条对标**：每个招标要求使用 `【招标要求】+【我方响应】` 模式
- **具体化**：避免空泛承诺，提供数字、时间、流程等可量化信息
- **表格化**：结构化信息优先使用表格呈现
- **评分导向**：根据评分权重分配内容详略

## 脚本工具

```bash
# 生成投标文件模板
conda run -n bid-skills python scripts/generate_bid_template.py --output <输出路径.docx>

# 提取招标 PDF 文本
conda run -n bid-skills python scripts/extract_pdf.py --input <招标文件.pdf> [--output <输出路径.txt>]

# 分析标书 docx 格式
conda run -n bid-skills python scripts/analyze_docx_format.py --input <标书示例.docx> [--output-dir <输出目录>]
```

## 参考文档

| 文件                                      | 说明                                   |
| ----------------------------------------- | -------------------------------------- |
| `references/bid_structure_knowledge.md` | 标书标准结构和各章节写作要点           |
| `references/scoring_strategy.md`        | 评分策略和得分技巧                     |
| `references/procurement_analysis.md`    | 招标文件解读方法论                     |
| `references/format_spec.md`             | 标书格式规范（字体、段落、表格、页面） |
| `references/writing_guide.md`           | 各章节的详细写作指南和范例             |

## 目录结构

```
.
├── SKILL.md              # 技能定义文件
├── scripts/
│   ├── generate_bid_template.py   # 模板生成脚本
│   ├── extract_pdf.py             # PDF 提取脚本
│   └── analyze_docx_format.py     # 格式分析脚本
└── references/
    ├── bid_structure_knowledge.md  # 结构知识
    ├── scoring_strategy.md         # 评分策略
    ├── procurement_analysis.md     # 解读方法论
    ├── format_spec.md              # 格式规范
    └── writing_guide.md            # 写作指南
```
