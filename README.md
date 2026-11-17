# 成绩分析与报告生成引擎（通用版）

> 🏷️ **E.D.STUDIO** · 成绩分析报告工具（通用版）


把一份标准列契约的成绩 Excel，一键变成 **Word 学科能力分析报告** + **交互式 HTML 网页版**。面向教师与班主任，零个人数据、配置驱动、可自由改造。

## 它能做什么

| 能力 | 说明 |
|---|---|
| 层次折算 | 按「校内排名」经可配置锚点折算为层次/院校档次，区间映射，附"经验估算"免责声明 |
| 四块分析 | 每个学科统一产出：能力摘要 → 学生分层 → 薄弱诊断 → 学科小结 |
| 双格式交付 | 一次运行同时产出 `.docx`（横向 A4、规范字体）与 `.html`（筛选/排序/搜索） |
| 个人报告联动（可选） | 交互 HTML 可内嵌 PDF 个人报告，点击学生直达对应页 |
| 完全通用 | 学校名、班级、科目、考试名、折算锚点、诊断归因全部由 `config.yaml` 注入 |

## 快速开始

```bash
# 1. 安装依赖
pip install openpyxl python-docx PyYAML

# 2. 复制配置并填写
cp config.example.yaml config.yaml
# 编辑 config.yaml：学校、班级、科目、折算锚点...

# 3. （可选）生成虚构示例数据先跑通
python scripts/make_sample.py --config config.yaml --out sample.xlsx

# 4. 运行
python scripts/build_report.py --config config.yaml --input sample.xlsx --out-dir output/

# 5. 对外分享前脱敏自检
python scripts/privacy_check.py --root .
# 严格模式：把专属词表放仓库外，用 --words-file 传入（词表本身不要进仓库）
python scripts/privacy_check.py --root . --strict --words-file /path/to/private_words.txt
```

产物在 `output/` 下：Word 报告 + 交互 HTML。

## 不会命令行？两种傻瓜化用法

除了上面的命令行方式，本项目还提供两条零门槛路径，适合分享给不懂技术的老师：

| 方式 | 文件 | 适用场景 |
|---|---|---|
| **本地一键包** | `一键运行_Mac.command` / `一键运行_Win.bat` | 双击启动 → 选 Excel → 自动识别列 → 生成。全程本机运行，成绩数据不出电脑，首次需联网装依赖，之后可离线 |
| **云端网页版** | `scripts/web_app.py`（Streamlit） | 浏览器打开 → 上传 Excel → 一键生成并下载。零安装，手机可用；数据仅在临时目录处理、会话结束即销毁 |

本地一键包会引导你填基本信息、自动识别 Excel 列，并生成 `config.yaml`，无需手写配置。
云端网页版可自行部署：`pip install -r requirements.txt` 后 `streamlit run scripts/web_app.py`。

完整转发文案、操作步骤与隐私对照见 `分享指南.md`。

## 文档

- `references/INPUT_SCHEMA.md` — Excel 列契约
- `references/CONFIG_GUIDE.md` — 配置字段逐项说明
- `references/CUSTOMIZE.md` — 折算规则与诊断归因的定制方法

## 隐私承诺

本仓库**不含任何真实个人数据**：所有示例均为虚构，配置模板使用占位符；扫描器本身也不内置任何真实姓名/学校词表（词表即个人数据，需放仓库外通过 `--words-file` 传入）。对外分享前请运行 `scripts/privacy_check.py`，它会扫描通用个人数据特征（邮箱、绝对路径、身份证、手机号、GitHub 账号等）；命中即非零退出。

> 注意：自动扫描是辅助手段，不能替代人工判断。发布前请人工复核 README、提交历史与所有示例文件，并用你自己的私有词表做严格扫描。

## License

MIT — 见 `LICENSE`。

---

本工具由 **E.D.STUDIO** 整理出品。
