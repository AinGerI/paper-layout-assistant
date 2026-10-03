# 两份原skill逐项继承对照

本次重新核对的来源是用户实际提供的：

- A：`10｜word排版Skills.zip`，同时检查“01 原始解压包”“02 修订版源码”“04 补充排版实现参考”及模板库说明。
- B：`paper-word-delivery/` 的SKILL及全部三份references。没有拿已安装的另一份skill代替该文件夹。

结论：本正式包将排版经验纳入主动阅读规则、细化经验、原始全文/实现和本对照表。原文保留是为了可追溯；AI执行时必须遵循当前主文和用户要求，不直接照跑历史脚本。

状态：“脚本落实”指现有JSON转换器/固定宏有对应行为；“条件流程”指已纳入必须遵循的处理和验收规则，按需求开展；“原引擎保留”指源代码及配置完整随包，须在适用Windows环境使用，不等于本次重新执行通过。

## A：word-format

| 原文位置/经验 | 新skill中的落点 | 状态与取舍 |
|---|---|---|
| SKILL Core rule：要求与引擎分离 | formatting + inherited-layout §1；JSON/YAML分流 | 脚本落实，不按论文重写宏 |
| Input priority：要求、模板、说明、未知项 | inherited-layout §1；legacy-word-format | 条件流程；既有Word未知保留，新稿采用明确默认 |
| inspect→纠偏→profile→副本→报告→validate | legacy-word-format及保留的scripts | 原引擎保留；不追加为每次新稿的额外步骤 |
| 段落序号与regex人工纠偏 | legacy-word-format末节 | 原引擎保留；段落号不跨稿套用 |
| 不覆盖、不改文字数字/公式、不转公式图片 | SKILL任务绑定；inherited-layout §1/8 | 默认保留；固定宏原位保存是后来明确要求的例外 |
| 不自动更新TOC/参考文献/交叉引用域 | inherited-layout §6；legacy-word-format | 保留为要求；固定宏自有遍历筛选MathType域，并在插件调用期间保护普通域、完成后恢复锁状态 |
| 公式默认避免固定行距；表行auto | inherited-layout §4/5 | 条件流程；新稿1.5倍不会固定裁高公式 |
| 浮动图只按明确要求改；正文story验收边界 | inherited-layout §6/8 | 条件流程；主文数量不冒充全story验证 |
| invalid keys/碰撞输出拒绝、失败副本不可交付 | 新脚本校验；legacy-word-format | 脚本落实/原引擎保留 |
| 角色：图段/图题/表题/摘要/关键词/文献标题 | legacy能力表；inherited-layout §3 | 部分JSON落实；特殊角色/局部字体走条件处理 |
| profile-guide：中文字号完整换算 | inherited-layout §3 | 完整保留 |
| single/1.5/double/exact/at_least/multiple | inherited-layout §4；legacy能力表 | JSON倍数；其余原引擎保留，不虚构字段 |
| 边框retain/none/three_line/full_grid | inherited-layout §5；legacy能力表 | 新稿无原格式retain意义；既有Word原引擎支持 |
| profile-sources：三种来源和extract局限 | inherited-layout §1/2；legacy模板节 | 原说明全文+提取器保留 |
| 全表/单元格水平/垂直对齐、内边距/行高 | inherited-layout §5；legacy能力表 | 原引擎保留，需目标稿页面复核 |
| OMath范围绕开普通字体设置 | inherited-layout §3/8；legacy对象保护 | 原代码保留；不以OMML数量充当MathType数量 |
| 示例profile不是正式规范；UTF8/BOM、schema验证 | legacy +原profiles/指南 | 原文件与检查器完整保留 |
| local-workspace：模板库/稿件/结果/维护分开 | inherited-layout §1；原模板库说明 | 保留组织方式，不沿用旧E盘/C盘路径 |
| 比赛/年份隔离、参考论文不是官方模板 | inherited-layout §1 | 保留并明确路由到模板库/分类说明；不复制论文私人内容当默认模板 |
| 源README：独立Word实例、保存后重开检查 | legacy-word-format | 原引擎保留；当前禁止未经允许操作桌面 |
| 补充：隶书多通道、关键词标签/内容分开 | inherited-layout §3；legacy补充节 | 条件流程+原补充代码保留 |
| 补充：0字符后清理残留磅缩进 | inherited-layout §2；新脚本角色缩进 | 新稿落实；既有稿按范围同步清理 |
| 补充：顶端对齐、黑色/阴影/缩放设置 | inherited-layout §3；补充配置 | 条件流程，仅明确要求时作用于文字 |
| 补充：数学原字体、已有中心/右tab | inherited-layout §3；legacy补充节 | 条件流程；不替换固定宏原生编号 |
| 补充：文档网格、命名空间/ignorable保护 | inherited-layout §4；legacy末节 | 保留；历史全篇网格设置被当前局部规则替代 |
| 补充：图片宽度正文上限/保比例、内容媒体比对 | inherited-layout §6/8 | 新稿比例/边界落实；已实现只缩小并记录DPI估计；既有Word按范围验证 |
| 完成需DOCX、中文MD/JSON报告及待检查项 | legacy + verification | 旧引擎保留；日常生成记录不混进论文、不重复附宏 |

## B：paper-word-delivery

| 原文位置/经验 | 新skill中的落点 | 状态与取舍 |
|---|---|---|
| SKILL Bind request：范围/版本/权限/格式/状态/验收 | SKILL任务绑定；inherited-layout §1 | 条件流程必做，已明确项不重复询问 |
| Translate requirements：节选、同行、缩进、去框/号 | SKILL任务绑定；inherited-layout §1/7/8 | 保留局部范围，不误删真表格 |
| Inspect format：defaults/style/direct/numbering/sections | inherited-layout §2 | 完整诊断经验保留 |
| 首行/悬挂/整段、字符/磅冲突 | inherited-layout §2/4 | 新稿字符设置落实；既有稿局部处理 |
| 软换行拉伸、keep规则、不要空格排版 | inherited-layout §4 | 条件流程+现样式实现 |
| Preserve content：嵌套结构、宏、图表、资产保真 | inherited-layout §5/6/8；原SKILL全文 | 条件流程；不宣称解析器已支持任意TeX/Word |
| Verify/Deliver：最终保存、内容/包/渲染/更改后复验 | inherited-layout §8；verification | 必须分结构验证与目标端结果，不交陈旧预览 |
| format: 非目录正文取样/摘录编号重启/页码零 | inherited-layout §2/4；原format全文 | 保留具体诊断线索 |
| format: 三线表层叠边框/列宽/表注/逐行/跨页 | inherited-layout §5 | 默认基础表已实现；复杂列宽等按需求处理 |
| format: bibliography数据顺序/前导空格/悬挂 | inherited-layout §7 | 保留；当前JSON非负首行不能代替悬挂排版 |
| format: 图片关系、crop/transform、TIFF透明度 | inherited-layout §6 | 保留；当前新稿图片等比转换实现 |
| format: 比较统计口径/实际页数/末页占用 | inherited-layout §7 | 条件流程；不得自动缩字删内容 |
| format: 替换边界/定位短语/纯净复制文本/联动修改 | inherited-layout §7 | 条件流程，不重建整篇 |
| format: 有序覆盖/未改部分/字体环境/刷新预览 | inherited-layout §8 | 条件流程；不只检查数量/关键词 |
| code: 摘录与可执行包、缩进注释语义、字体选择 | inherited-layout §7；原code全文 | 条件流程；代码围栏需专门处理，当前解析器拒绝 |
| code: 可编辑高亮/长行/Tab/不插弯引号软连字符 | inherited-layout §7 | 完整保留 |
| code: 题注、真实标题、字段和跨引用一致 | inherited-layout §4/6；macro-use | 保留；图表自动交叉引用不是现脚本已有功能 |
| code: 紧凑行距不裁切/长表重复表头/局部范围 | inherited-layout §4/5/7 | 基础实现+条件验收 |
| equations: TeX/OMML/OLE/图四态、外观与可编辑性 | inherited-layout §8 | 辨识知识保留；新稿默认TeX路径不改 |
| equations: 转义、已知宏、外框、多行/引用/裁切 | inherited-layout §7/8 | 保留，公式源码保持语义 |
| equations: 对齐目标/矩阵列宽/版本/导入不同机制 | inherited-layout §8；原equations全文 | 仅明确要求对齐时核验；不恢复强制对齐 |
| equations: 失败路线/不循环猜语法/三个官方链接 | inherited-layout §8；原equations全文 | 保留版本核实要求，不将历史失败全局禁用 |
| equations: OLE/MTEF恢复/源序预览/高风险字形验收 | inherited-layout §8 | 条件流程；不声称已有全自动恢复工具 |
| equations: 单列array例子与中文TeX/行距兼容线索 | 原equations全文；markdown-contract | 保留候选，不声称所有版本通用 |

## 完整原文与代码的保留范围

[文件清单](inheritance-manifest.json)逐个记录42份源文件与包内位置，复制时比较原始字节一致。源SKILL文件改名为SOURCE_SKILL.md以免被误当另一个自动发现入口，内容不变。

- B的主文与3份参考说明完整放入references/original-paper-word-delivery。
- A的原版与修订版主文、README、profiles、references、全部Python脚本与requirements完整保留在assets相应目录。日常以修订版为参考，旧版用于追溯。
- A补充实现的3份Python参考与2份YAML配置完整保留。
- 模板库使用说明及参考论文的已提取YAML配置保留；参考论文是实例，不作为所有稿件模板。

未作为新能力带入：历史安装记录、历史测试报告、清理脚本、自动安装PowerShell、机器发现元数据、参考论文原稿及机器采样JSON。这些是旧环境操作/材料记录，不是遗漏排版规范；原ZIP保持原样可查。旧环境“已安装/已测试”的结论不能转移到当前包。

## 保持的当前用户规则

用户通过的PWDMathType133Check宏逐字节冻结；对象缩放提示继续取消；不强求等号对齐；正文默认2字符、题注正文同字体且6/6磅；只修公式段网格；宏按需备份并原位保存；日常只提供Markdown/Word，首次或更新才提供通用宏；未经允许不Computer Use/启动桌面自动化。

最终复核补充：原引擎正文story快照边界已显式写入日常验收；题注前缀/邻接启发式的误判已通过独立前向用例确认，新增强制语义核对与局部角色纠偏，不宣称解析器已自动消除歧义。
