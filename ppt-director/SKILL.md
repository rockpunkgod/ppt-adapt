---
name: ppt-director
description: 为整套或单页PPT制定叙事、全局视觉系统与页面计划，输出可校验的设计规格和执行契约。用于制作前总控与改稿规划；具体OfficeCLI制作和动效验证交给执行层。
---

# PPT Director

先决定全稿讲什么、观众应理解或决定什么，再决定每页使用什么表现形式。沿用用户已给出的受众、素材、品牌和制作授权；不要重新索要确认，也不强制用户参与三套方案评审。

艺术方向与整套创意变化遵循[制作总则](../officecli-ppt-design/references/deck-design.md)：开放创意任务主动探索，方向由内容与受众决定；随机联想只帮助打破惯性。先从内容关系提出视觉主体、空间隐喻与观看顺序，再选已有配方实现。开放需求比较两到三个封面/代表内容页方向，按理解效果与整套展开潜力选择；品牌续稿与局部改稿保留必要约束。全稿同时规划页面角色、构图尺度、阅读视角与动静节奏，不能只有封面有个性。以真实预览和减法精修完善作品，复用机制而不复用固定美术模板。候选与关键取舍记入已有理由文档即可，不增设审美评分或契约字段。

相邻页同时规划注意力交接与顺滑衔接，说明共享场景、必要停顿及页内动作的先后。变化的构图保持连贯的运动语法，避免无意硬切和入场前的空白等待；制作后以实际连续播放确认。

使用前读 [契约与门控](references/contracts.md)，确定方向时参考[全稿设计决策](references/design-direction.md)。输出 `director-rationale.md`、`design-spec.json`、`deck-plan.json`。理由文档简述叙事选择、视觉规则、页面节奏、变体原因及证据缺口；机器JSON是执行的唯一权威，理由文档不能暗中改写它。

先形成全局spec，再规划页面：

- 叙事要有 opening、throughline、closing_ask；单页请求同样写清该页的开场语境、核心论点与下一步，不强行扩成多页。
- 每页写一个明确message和page job（`role`承担，必要时补`page_job`），指定证据、信息关系和layout family。证据可以是用户原文、数据或明确标记的原创演示内容，不能伪装成已核实事实。
- 字体、字号、色彩、间距来自spec。保持共同视觉语法，但根据语义使用图表、时间轴、对比、章节或留白；不要把整套稿排成同款卡片墙。
- 页角色可采用 opening/context/evidence/comparison/process/decision/closing/breathing 等语义词，无固定页数或角色比例。连续多页同角色可以合理；解释连续节奏或单一layout的原因。
- 深色、出版式等局部变体必须在spec声明页范围和原因；不得临时逐页改色。对页内文字设置明确字体角色，重要信息不得缩到脚注字号来塞内容。
- 每个可编辑对象有稳定id、实际厘米bounds；跨页连续对象可用morph_key，不能给同页两个对象同一key。图表/表格的真实数据引用要可追溯。
- 只规划用户目标需要的机制。执行层按现有已验能力选择OfficeCLI、COM、混合或必要外部渲染；某配方代表案例已过不代表任意参数/素材都过。计划不声称制作或放映已经通过。

运行离线检查：

```text
python scripts/validate_deck_plan.py --spec /project/design-spec.json --plan /project/deck-plan.json
python scripts/validate_deck_plan.py --spec /project/design-spec.json --plan /project/deck-plan.json --page s05 --report /project/check-s05.json
```

修复错误再交给组合/制作层。警告需要判断或写明理由，不能机械增加页数、配色或动效消除警告。校验只证明契约一致性：不检测实际字体安装、文字溢出、对比度、图表编辑能力或播放效果；这些必须由制作后的原生检查完成。

交接时给出三份文件路径与校验结果、缺失素材、待验机制及不可改的事实/字面约束。保持既有文件原件；更新现有计划要说明影响页，不把一页改稿扩大为整套重新设计。不导入参考项目SVG转换引擎、工作区系统或供应商配置。
