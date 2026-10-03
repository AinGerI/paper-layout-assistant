# 原 word-format 排版引擎与补充经验的调用边界

## 两条路径，各用自己的配置

日常路径保持：规范Markdown → 本包scripts/markdown_to_word.py + JSON profile → 用户在Windows运行固定宏。

已有Word保真重排、正式模板取样，或JSON尚不能表达的高级表格/行距设置，按本页处理。包内 `assets/word-format-revised/` 保存用户原ZIP中的修订版排版引擎、YAML模板和完整说明；`assets/word-format-original/` 保存更早原版，供查差异，不作为默认执行版。原始文件未修改。

| 项目 | 新稿JSON路径 | 原Word YAML路径 |
|---|---|---|
| 输入 | 规范Markdown | 既有DOCX |
| 未知格式 | 使用已说明的默认值 | null/disabled保留原样 |
| 排版工具 | 跨平台文件转换，不启动Word | Windows桌面Word COM |
| 角色 | title、三级heading、body、caption、equation、table、references | 摘要标题、关键词、图段、图题、表题、参考文献标题等独立角色及人工识别覆盖 |
| 行距 | 正数倍数 | 单/1.5/双倍、multiple、exact、at_least |
| 表格 | 三线/全网格/无框、重复表头、跨页、等宽初始布局 | 整表位置、宽度、自动适应、文字对齐、垂直对齐、内边距、行高、分线宽度 |
| 页面 | 纸张宽高及四边距 | 加横竖方向、装订线、页眉页脚距离 |
| 输出 | 新DOCX+转换记录 | 排版副本+中文MD/JSON报告+保存后比较 |

不能把两种schema_version=1当作同一个schema；YAML不能直接交给JSON转换器。JSON不存在的字段应选择专门处理或扩展测试，不能忽略后宣称满足。用户没有需求时不额外跑一轮旧引擎。

## 模板/规范读取

先读 [源profile指南](../assets/word-format-revised/references/profile-guide.md)、[规范来源流程](../assets/word-format-revised/references/profile-sources.md) 和 [源YAML模板](../assets/word-format-revised/profiles/profile-template.yaml)。记录文字规范来源；官方模板与参考论文分开。

用户指定模板库或比赛类别时，先读当前指定模板库的总说明和对应分类说明，核对该分类的特定补充步骤。使用包内历史示例时，参照 [模板库说明](original-template-library/01%20模板库/模板库说明.md) 和 [华为杯分类说明](original-template-library/01%20模板库/华为杯/使用说明.md)；只迁移已核对规则，不沿用历史路径或将参考论文当作官方模板。

用户允许在Windows本地运行时，原脚本支持以下顺序：check_profile → inspect_word → 纠正检测角色 → format_word（新路径）→ validate_word（保存后重新打开）。extract_profile 可先从模板提取起点。脚本用独立Word实例，不附着用户当前打开的Word；失败副本和失败报告不能算完成。

以任务副本中的Python环境和脚本绝对路径调用，示例变量都由当前任务实际位置填写：

```powershell
$taskPython = 'D:\当前任务\.venv\Scripts\python.exe'
$engineDir = 'D:\当前任务\markdown-word-mathtype\assets\word-format-revised'
& $taskPython "$engineDir\scripts\check_profile.py" 'D:\当前任务\目标.yaml'
& $taskPython "$engineDir\scripts\inspect_word.py" 'D:\当前任务\原稿.docx' --profile 'D:\当前任务\目标.yaml'
& $taskPython "$engineDir\scripts\format_word.py" 'D:\当前任务\原稿.docx' --profile 'D:\当前任务\目标.yaml' --output 'D:\当前任务\排版副本.docx'
& $taskPython "$engineDir\scripts\validate_word.py" 'D:\当前任务\原稿.docx' 'D:\当前任务\排版副本.docx'
```

所需pywin32/PyYAML见原requirements.txt。这是可选Windows流程，不加入日常跨平台环境，不自动安装、不运行历史安装脚本。当前用户未允许Computer Use或桌面自动化；AI不能因为源skill写着“运行COM”就启动Word。由用户在Windows执行，或获得明确的该项授权后执行。Mac不尝试替代自动操作Word。

## 三个不能漏掉的区别

1. **现有对象保护。** 原 `apply_figures` 遍历全部InlineShapes，未按真实图片与OLE公式细分；`document_snapshot.equations`计数是OMML，不是完整MathType计数。因此对已有MathType的稿件，必须关闭原引擎figure设置，另行针对真实图片处理；增加OLE身份/数量与预览核对。原引擎完整性快照以正文main story为边界；页眉、页脚、文本框等若在任务范围内需另行核验，不以正文计数代表全文。不能把旧引擎直接当固定宏的无条件后处理。
2. **保存机制不同。** 原引擎保留原稿、生成排版副本；固定宏是用户要求的可选备份后原位保存。不要让原引擎的“永远另存副本”覆盖宏的已测试保存流程。仅在用户要既有稿重排时启用原引擎。
3. **域的更新范围。** 原引擎仅重新分页，不自动更新目录/交叉引用/文献域。固定宏的自有遍历只进入MTPlaceRef、MTEditEquationSection及其数字后缀包装域，更新精确命名的MTEqn/MTChap/MTSec计数器；普通SEQ/MACROBUTTON不更新，也不进入其内部递归。此外，插件调用期间临时锁定无关域并恢复原状态，覆盖所有story；使用与验证见 [宏说明](macro-use.md)。

## 人工纠偏和补充实现

原检测有段落序号paragraph_overrides与regex_overrides。根据inspect结果修正标题/正文/公式/题注误分类，不能只凭字号或“图1”开头确定角色。段落序号随文档编辑会变，不能跨稿套用旧序号。

`assets/word-format-supplement/` 保留原“补充排版实现参考”的代码与配置。这里包含：关键词标签/内容分别设置、隶书多字体通道、清理字符/磅缩进、页面顶端对齐、数学字体保护、已有中心/右制表位、图片正文宽度上限、节网格保留、保存后内容/图片/公式比对。这些是实现依据，不是通用可直接执行的脚本：历史路径、段落序号、样稿和测试断言都绑定旧任务。

按当前任务移植某项时：选取对应局部处理 → 绑定当前范围与实际模板值 → 在副本上运行 → 比较未改内容和对象 → 保存后重新读取及渲染。涉及Word必须遵守当前执行授权。历史补充的全局字体/全段网格或固定False/True设置不能覆盖用户现在的“只修公式段落网格”规则；保留固定宏的校准及对象定位。

现成公式居中/编号靠右如依赖制表符，排版时只调整已有制表位，不插入文本或改编号；当前新稿则继续使用MathType原生编号，不把旧制表位排法替换进宏。

OOXML补丁应保留命名空间前缀及mc:Ignorable绑定，避免简单序列化丢声明；节网格从指定模板按节复制时先确认该模板是本次要求。旧例子“所有节用同一网格”不是多节论文的通用规则。
