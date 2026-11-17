---
name: score-analysis-report-engine
description: 把一份标准化成绩 Excel 转换为「学科能力与分层分析」Word 报告 + 交互式 HTML 的通用引擎。适用于高中/初中班级的多科成绩分析：按校内排名折算层次、逐生能力诊断、生成分学科表格与可筛选排序的网页。触发词：成绩分析报告、学情分析、能力分析表、成绩交互 HTML、批量生成分析报告、院校层次折算。
---

# 成绩分析与报告生成引擎（通用版）

## 用途

输入一份**标准列契约**的成绩 Excel 与一份 **config.yaml**，输出两份产物：

1. **Word 报告**（`.docx`）：横跨多学科的「能力摘要 → 学生分层 → 薄弱诊断 → 学科小结」四块表格，附层次/院校档次说明。
2. **交互 HTML**（`.html`）：同口径数据的网页版，支持层次下拉筛选、姓名搜索、点击列头排序；可选内嵌 PDF 个人报告。

所有学校名、班级名、学生名单、考试名、折算锚点均由使用者提供，本 skill **不含任何真实个人数据**。

## 何时使用

- 教师拿到班级多科成绩，需要快速产出结构化分析报告
- 需要把同一批数据同时交付 Word（打印/存档）与 HTML（交互查看）
- 需要按「校内排名 → 层次/院校档次」做经验折算与逐生分层

##  workflow（4 步）

### 1. 准备配置
复制 `config.example.yaml` 为 `config.yaml`，按 `references/CONFIG_GUIDE.md` 填写学校、班级、考试科目、折算锚点、诊断规则。
**必填项**：`school`、`class_name`、`subjects`、`tier_rules`、`input.columns`。
**可选项**：`pdf_report`（个人报告 PDF 模板与页码规则）、`school_recommend`（层次对应院校）。

### 2. 准备数据
按 `references/INPUT_SCHEMA.md` 整理 Excel。列名须与 `config.yaml → input.columns` 一致。
可用 `scripts/make_sample.py` 生成一份**全虚构**示例数据先跑通流程。

### 3. 运行引擎

```bash
python scripts/build_report.py --config config.yaml --input data.xlsx --out-dir output/
```

生成：
- `output/<class_name>_学科能力分析.docx`
- `output/<class_name>_学科能力分析（交互版）.html`

### 4. 脱敏自检（对外分享前必做）

```bash
# 通用特征扫描（邮箱、绝对路径、手机号、身份证等）
python scripts/privacy_check.py --root .

# 严格模式：追加你自己的专属词表（文件放仓库外，不进包）
python scripts/privacy_check.py --root . --strict --words-file /path/to/private_words.txt
```

`private_words.txt` 每行一个真实姓名/学校/班级/考试名，**不要放进仓库**——词表本身也是个人数据。扫描脚本、文档、配置模板、git 历史；**0 命中**才算合格。

## 关键约定

- **折算说明**：校内排名是校内坐标系，折算出的全省位次/院校层次为经验估算，文档须保留「仅供预判参考」字样，不可当精确预测。
- **数据口径**：`期初 → 期中` 两考为默认链条；如只有一次考试，在 config 中设 `exam_chain: [midterm]`。
- **排版规范**：Word 中文华文楷体、英文 Times New Roman，10.5pt，1.5 倍行距，横向 A4 表格。
- **视觉规范**：HTML 默认薄荷绿 + 白、圆角卡片、大按钮、触屏友好。

## 目录结构

```
score-analysis-report-engine/
├── SKILL.md                     # 本文件
├── README.md                    # 对外说明
├── LICENSE                      # MIT
├── config.example.yaml          # 全虚构配置模板
├── scripts/
│   ├── build_report.py          # 主引擎
│   ├── make_sample.py           # 全虚构示例数据生成器
│   └── privacy_check.py         # 脱敏门禁
├── references/
│   ├── INPUT_SCHEMA.md          # 输入 Excel 列契约
│   ├── CONFIG_GUIDE.md          # 配置字段说明
│   └── CUSTOMIZE.md             # 折现/诊断规则定制
└── assets/                      # （空）使用者自有数据不入库
```

## 依赖

Python 3.9+，`openpyxl`、`python-docx`、`PyYAML`。
交互 HTML 为单文件，无前端依赖；若启用 PDF 内嵌，需把 PDF 放在 HTML 同目录。

---

本工具由 **E.D.STUDIO** 整理出品。
