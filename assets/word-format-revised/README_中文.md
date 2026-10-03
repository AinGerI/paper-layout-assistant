## 本机日常使用入口

工作区：`C:\paper-workspace`。

- `01 模板库`：按华为杯、统计建模等分类存放模板与要求。
- `02 待排版论文`：放论文原稿。
- `03 排版结果`：保存排版副本和中文报告。
- `90 维护备份`：程序与历史记录，平时无需打开。

直接告诉 Codex：“使用 word-format，按模板库里的华为杯要求排版这份论文。”

现有 reference-paper.docx 是参考论文，不是已核实的官方模板；具体限制请看华为杯分类的使用说明。
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

本机优先安装到用户指定的 `~/.codex/skills/word-format`，并通过 Codex runtime 的 `skills/list` 验证。当前官方文档也列出 `~/.agents/skills` 和项目 `.agents/skills`；不要仅凭目录存在就声称已识别。

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


## 本机修订版

- 使用 Skill 目录内 `.venv/Scripts/python.exe`，命令和输入都使用绝对路径。
- 脚本通过独立 Word 实例处理副本，拒绝覆盖原稿和已有输出；报错不保存未完成的内存修改。
- 保存后重新打开检查正文文本及公式、表格、InlineShapes、Shapes、Fields 数量。
- Word COM 需要桌面会话权限；Codex 沙箱内启动失败时需通过执行权限审核运行。
- `scripts/check_profile.py` 校验 Profile；`scripts/smoke_test.py --output-dir <新的测试目录>` 执行真实 Word 测试。
- 三种 Profile 来源及能力边界见 `references/profile-sources.md`。PDF/Word 说明由 Codex 阅读整理，尚无全自动解析器。
- 不自动更新域、不重建公式。完整性检查主要覆盖正文 story；复杂浮动布局、页眉页脚和特殊编号仍需人工查看。
- 原下载 ZIP 保留；原解压包与修订源码分别存放在工作区“90 维护备份”的“01 原始解压包”和“02 修订版源码”。旧测试样稿、临时报告和预览文件已按用户要求清理。

