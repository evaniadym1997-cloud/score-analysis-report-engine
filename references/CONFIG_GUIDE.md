# 配置字段说明

`config.yaml` 共 9 个顶层字段。带「必填」的字段缺失会导致引擎报错退出。

| 字段 | 必填 | 类型 | 说明 |
|---|---|---|---|
| `school` | 是 | 字符串 | 学校显示名，出现在报告标题/说明。用你自己的学校名或化名。 |
| `class_name` | 是 | 字符串 | 班级显示名，用于产物文件名与标题。 |
| `term` | 否 | 字符串 | 学期描述，出现在副标题。 |
| `update_date` | 否 | 字符串 | 报告日期；建议 `YYYY-MM-DD`。 |
| `analyst` | 否 | 字符串 | 署名/来源。 |
| `exam_chain` | 是 | 列表 | 考试名称列表，约定为从早到晚。1 项或 2 项。 |
| `subjects` | 是 | 字典 | 学科 → `full`（满分）、`sections`（板块 → 满分）。 |
| `input` | 是 | 字典 | `columns`（列映射），详见 INPUT_SCHEMA.md。 |
| `tier_rules` | 是 | 列表 | 折算规则，详见 CUSTOMIZE.md。 |
| `missing_tier` | 否 | 字符串 | 排名缺失时的标签，默认「数据不全」。 |
| `school_recommend` | 否 | 字典 | 层次 → 院校列表；不需要可删除。 |
| `weak_reasons` | 否 | 字典 | 板块 → `{cause, advice}`；未命中用 `default_reason`。 |
| `pdf_report` | 否 | 字典 | `enabled: true` 时 HTML 增加「查看报告」列；PDF 需自行放入 HTML 同目录。 |

## 折算规则 `tier_rules`

按列表顺序从上到下匹配，命中第一条即停止：

```yaml
tier_rules:
  - max: 50        # 排名 <= 50
    tier: "拔尖层"
    color: "#FCE4D6"
    desc: "说明文字"
  - max: null      # null = 兜底，必须放最后
    tier: "攻坚层"
```

- `max` 必须是数字或 `null`；引擎按列表顺序判断，不自动排序，所以请**从小到大**写。
- `color` 是 HTML 中该层次标签的背景色（十六进制）。
- `desc` 会写入 Word 的层次说明表。
- 折算说明中的「全省位次/院校层次」为经验估算，引擎会自动附「仅供参考」免责声明；请勿删除。

## 层次与院校推荐

`school_recommend` 的 key 必须与某个 `tier` 完全一致（大小写、空格都要一致），否则该层次不显示院校列。

## 考试名称一致性

`exam_chain` 里的每一项（如「期初」「期中」）会：

1. 作为 `input.columns.totals.<学科>` 下的键；
2. 出现在报告表头（如「语文期初」「语文期中」）；
3. 用于计算「较期初增减」。

改考试名时，这三处都要同步改。

## 常见错误

| 现象 | 原因 |
|---|---|
| 报表全为 `-` | Excel 列名与配置不一致，或表头不在第 1 行 |
| 所有学生都是 `missing_tier` | `rank` 列映射错误，或排名列存的是文本 |
| 薄弱诊断显示「未配置」 | 未填 `section_rates`，或板块名与 `subjects.sections` 不一致 |
| HTML 无「查看报告」列 | `pdf_report.enabled` 为 false，或 PDF 未放同目录 |

---

本工具由 **E.D.STUDIO** 整理出品。
