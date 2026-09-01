# hit-thesis-proposal-format

把哈尔滨工业大学（HIT）硕士学位开题报告草稿一键调整为符合《硕士学位论文报告模板（2025）》规范格式的 Agent Skill。

本 Skill 只改格式、不改正文：输入任意内容完整的开题报告草稿（正文 + 层级标题 + 参考文献，可有封面），输出套用模板格式的新 docx。

## 功能特性

- **页面设置**：A4，上 2.5 / 下 2.2 / 左 2.5 / 右 2.5cm，无页眉
- **字体**：正文宋体小四（12pt），各级标题黑体，数字/英文 Times New Roman
- **层级标题**：节（15pt，段前/段后 0.5 行）、条（14pt，段前/段后 0.5 行）、款（12pt），均顶格
- **固定行距**：21pt（每页约 33 行）
- **正文缩进**：首行缩进 2 字符；`（1）（2）…` 开头的项按正文处理（题序空 4 个半角字符）
- **参考文献**：宋体/TNR 12pt，悬挂缩进（首行顶格、续行缩进 2 字符）
- **封面**：可自动识别已有封面并套用格式，也可用 `--cover-json` 重建模板封面；居中行无缩进，填表行下划线通过制表位对齐
- **不改变正文文字**：仅调整格式，内容原样保留

## 目录结构

```
hit-thesis-proposal-format/
├── SKILL.md                        # Skill 入口与使用说明
├── scripts/
│   └── format_proposal.py          # 格式调整脚本（python-docx）
├── references/
│   └── format-spec.md              # 完整格式规范（模板逐项提取）
└── assets/
    └── 硕士学位论文报告模板（2025）.docx   # 模板原件（未随仓库同步，见下方说明）
```

## 环境要求

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/)：`pip install python-docx`

## 安装（作为 Agent Skill）

把 `hit-thesis-proposal-format` 整个文件夹放入 Agent 的 Skill 目录（路径以 `workspace/.user_skills` 结尾）：

```
<环境对应父目录>/workspace/.user_skills/hit-thesis-proposal-format/
```

然后在对话中提供开题报告草稿并说明“按模板调整格式”即可触发。

## 使用方法

```bash
python scripts/format_proposal.py <输入.docx> <输出.docx> [--cover-json 封面.json] [--dry-run]
```

### 重建封面（可选）

```bash
python scripts/format_proposal.py draft.docx out.docx --cover-json cover.json
```

`cover.json` 字段：

```json
{
  "title": "论文题目",
  "degree": "学位论文",              // 或 "实践成果"
  "college": "土木工程学院",
  "discipline": "土木工程",
  "advisor": "导师姓名",
  "student": "研究生姓名",
  "student_id": "学号",
  "date": "2026年9月"
}
```

> 题目、姓名等无法从上下文确定时使用占位符，不要编造。若草稿已自带封面，脚本会自动识别并套用封面格式，无需 `--cover-json`。

## 格式规范摘要

完整规范见 [`references/format-spec.md`](references/format-spec.md)。核心参数：

| 项 | 值 |
|---|---|
| 页面 | A4，上 2.5 / 下 2.2 / 左 2.5 / 右 2.5cm，无页眉 |
| 正文 | 宋体 小四(12pt)，数字/英文 Times New Roman，固定行距 21pt，段前/段后 0，首行缩进 2 字符 |
| 节标题 | 黑体 15pt，段前/段后 0.5 行，顶格 |
| 条标题 | 黑体 14pt，段前/段后 0.5 行，顶格 |
| 款标题 | 黑体 12pt，段前/段后 0，顶格 |
| 项（（1）…） | 按正文处理（宋体小四，首行缩进 2 字符 = 题序空 4 个半角字符） |
| 参考文献 | 宋体/TNR 12pt，固定 21pt，悬挂缩进 |
| 封面居中行 | 加粗、居中、无缩进 |
| 封面填表行 | 标签无下划线，填充区（值）下划线，制表位对齐 |

## 校验产物

生成后用 python-docx 回读输出，抽查：

- `w:rFonts`（宋体/黑体/Times New Roman）与 `w:sz`（15/14/12pt）
- `w:spacing`（line=420, lineRule=exact）
- `w:ind`（正文 firstLineChars=200；标题顶格；参考文献 left=480+hanging=480；封面填表行 left=1260 + 制表位 4500）
- `w:snapToGrid=0`
- 用 LibreOffice 渲染 PDF 或逐页查看，检查封面居中/下划线对齐、标题层级、正文缩进、参考文献悬挂缩进、页眉为空
- 比对文本，确认正文文字未被改动

## 已知边界

- 行距固定 21pt 为模板建议；如需调整，修改脚本常量 `LINE_TWIPS` 后重跑。
- 脚本只清空页眉文字，不删除页眉链接关系。
- **模板原件（assets 下的 docx）未随本仓库同步**：GitHub API 无法经文本接口推送二进制文件。如需模板原件，请从哈工大研究生院获取，或使用网页端/本地 git 补充上传。

## 仓库内容与本地 Skill 的同步

仓库文本文件（`SKILL.md`、`scripts/`、`references/`）与本地安装的 Skill 保持同步；`assets/` 模板文件仅存在于本地安装目录。
