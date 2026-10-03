# word-format Profile 指南

Profile 只描述“目标格式”，不写论文内容。未知的规则保持 `null` 或 `enabled: false`，不要猜。

常用字号换算（pt）：初号 42，小初 36，一号 26，小一 24，二号 22，小二 18，三号 16，小三 15，四号 14，小四 12，五号 10.5，小五 9，六号 7.5，小六 6.5。

行距推荐写法：`single`、`1.5`、`double`；固定值用 `{rule: exact, value: 20}`；“最小值”用 `{rule: at_least, value: 18}`。公式段落优先使用 `at_least`，避免分式、矩阵被固定行距裁切。

表格边框 `type` 支持：`retain`（保留原样）、`none`、`three_line`、`full_grid`。三线表会清除竖线及内部横线，仅保留顶线、表头下线、底线。

`detection.paragraph_overrides` 用 Word 段落序号做最可靠的人工纠偏，例如：

```yaml
detection:
  paragraph_overrides:
    "1": title
    "18": heading_1
    "46": equation_block
```

`regex_overrides` 可按文字匹配：

```yaml
detection:
  regex_overrides:
    - pattern: '^附录\s*[A-Z]'
      role: heading_1
```

从官方模板提取时，脚本主要读取页面设置和 Word 内置的 Normal/Title/Heading/Caption 样式。模板若大量依赖手工直接格式而不是样式，必须再结合官方文字规范或检查报告补充 Profile。
