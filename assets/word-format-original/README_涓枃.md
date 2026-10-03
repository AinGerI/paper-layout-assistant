# word-format：通用 Microsoft Word 学术排版 Skill

这是一个面向 **Windows + 桌面版 Microsoft Word + Codex** 的通用排版 Skill。它不是华为杯专用工具：比赛、毕业论文、期刊和自定义格式都通过不同 YAML Profile 切换。

## 推荐安装：用户级 Skill

解压后，在这个文件夹里打开 PowerShell：

```powershell
powershell -ExecutionPolicy Bypass -File .\install_skill_to_codex.ps1 -Scope user
```

它会复制到：

```text
%USERPROFILE%\.codex\skills\word-format
```

然后进入该目录安装 Windows 依赖：

```powershell
cd $HOME\.codex\skills\word-format
powershell -ExecutionPolicy Bypass -File .\install_windows.ps1
```

如果只想让某一个 Codex 项目使用，可在该项目目录运行：

```powershell
powershell -ExecutionPolicy Bypass -File .\install_skill_to_codex.ps1 -Scope project -ProjectPath "你的项目目录"
```

Codex 官方支持仓库级 `.codex/skills/<skill-name>/SKILL.md` 和用户级 `~/.codex/skills/<skill-name>/SKILL.md`。

## 最简单的使用流程

1. 把论文 `.docx` 和格式要求/官方模板放到项目里。
2. 对 Codex 说“调用 word-format”。
3. 它先分析结构，再创建/提取 Profile，然后排版副本并做完整性检查。

推荐提示词：

> 调用 word-format。根据“官方格式要求.pdf”和“官方模板.docx”创建一个新的格式 Profile，先分析“论文.docx”的结构，确认标题、正文、公式、图题和表题识别没有明显问题后，再排版到“论文_目标格式.docx”。不要覆盖原稿，不要修改正文内容、数字或公式含义。完成后给我排版报告和仍需人工检查的问题。

## Profile 的三种来源

- **官方 Word 模板**：先用 `extract_profile.py` 提取页面和内置样式，再由 Codex 对照规范补齐。
- **PDF/Word 格式要求**：Codex读取明确规则，写入新 Profile。
- **自然语言要求**：例如“正文小四宋体，1.5 倍行距，三线表”，直接生成 Profile。

## 注意

`profiles/demo-chinese-academic.yaml` 只是功能演示，不代表任何比赛、学校或期刊规范。真正提交前必须用对应官方要求创建自己的 Profile。

当前 v0.1 的定位是“安全的通用基础版”：正文/标题/公式段落/表图题/表格/页边距/图片宽度已经可配置；复杂浮动图片、非常规合并表、特殊公式编号会主动列为人工检查项，而不是强行修改。
