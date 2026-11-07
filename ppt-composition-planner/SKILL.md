---
name: ppt-composition-planner
description: 将整套PPT设计规格和逐页计划组合成可执行依赖计划，选择已有OfficeCLI配方，检查样式、空间、Morph和时间轴冲突，并绑定整套原生验收。适用于多个已验机制一起使用；不负责渲染或重写单项配方。
---

# PPT组合规划

读取任务的设计规格、完整逐页计划和实际配方索引。使用现有 `officecli-ppt-design` 完成制作，本Skill不复制其制作引擎，也不执行PowerPoint、浏览器或外部服务。

保留[制作总则](../officecli-ppt-design/references/deck-design.md)确定的独立艺术方向与整套创意变化。配方提供对象机制、坐标与执行方法，不把示例配色、字体、卡片结构或动画顺序当成默认美术。组合应落实各页不同的表达任务与连续关系；有意的变化不能被统一样式抹平，也不为增加技巧数量追加机制。创意取舍先在全稿计划中确定，再绑定制作参数，保持既有契约和技术验收标准。

组合时间轴时保证相邻页与页内动作顺滑衔接，处理突兀启停、多余点击、空白等待与Morph误配。保留既有对象身份和必要的讲述停顿，用真实连续播放检查中间态，不因单项动画通过就认定整体过渡顺滑。

先读[冲突与输入契约](references/conflicts.md)。优先级固定为：内容完整正确 → 可读性 → 全局一致 → 动画 → 装饰。保留全部原文与数据；不能删内容、缩字号或把对象栅格化来假装修好冲突。有意的章节变化应在 `variants` 中声明作用页，不能被全局一致性检查抹掉。

运行标准库规划器，输出目录必须全新：

```powershell
python scripts/plan_composition.py --spec design-spec.json --deck-plan deck-plan.json --index ../officecli-ppt-design/references/recipe-index.md --out NEWDIR
```

它实际读取索引并绑定SHA，按每页 `required_mechanisms` 选择现有表格配方。优先使用索引中原子机制全文或完整脚本通配入口（例如 `theme_semantics_*`）；多行同用一脚本时明确机制全文或 `index-line-N`，不要猜一个英文别名就是新能力。保留配方原有后端和限制，内容/素材替换后不外推通过范围。

检查 `composition-plan.json` 的操作DAG、空间关系、属性write sets与毫秒/COM秒映射；阅读 `conflicts.json` 中真实自动修复的before/after/reason/priority，缺省token赋值另列 `defaults_applied`。每个选中配方补明确 `reuse_scope/excluded_capabilities`，不将原语复用冒称整配方。退出码2表示仍有阻断问题，应修输入或提供明确适配，不执行制作。规划器不证明文字适合框、字体实际可用或动画视觉正确，这些由整套原生验收负责。

交给制作层时要求：OfficeCLI基础 → CLI close/flush → 必要COM精修；同一文稿单写者，COM/UI独占并只关闭自己的稿。DAG列的是交接约束，不是可以直接运行的任意shell命令。每页按内容角色选版式；不要让全局ApplyToAll覆盖局部Morph/Zoom，也不要让继承动画和本地效果抢同一对象。

制作入口必须消费effective对象的 `native_name` 与 `font_size_pt`，不另造名称/字号。执行前调用脚本内 `verify_source_files(plan)`，核所有已绑定本地证据/输入/索引真实字节；改过data.json须重新规划，不能只因设计JSON没变就沿用旧hash。

首次输出始终 `planned`（或 `blocked`），且组合 `unverified`。已有配方的通过不能推出本次组合通过。必须使用当前计划和成品SHA绑定的**全套新原生**报告，才可另生成verified计划：

```powershell
python scripts/plan_composition.py --spec design-spec.json --deck-plan deck-plan.json --index ../officecli-ppt-design/references/recipe-index.md --out NEWVERIFIEDDIR --native-report native-report.json --deck final.pptx
```

报告格式见参考。需逐页原生PNG；含动态的组合还需实际放映录屏、点击/返回/停止等相关行为检查。报告必须是真实验收记录，不得为通过哈希检查编造报告。哈希只能绑定证据，不能证明录像来源。未匹配或未完成的验收保持未验证。

本Skill取自已研究的故事板与母版教程：先检查全套论点，再组合页面角色、样式和动效。来源与组合边界在参考末尾；这些原则不强制固定页数、指定字体或任意数值的动画密度。
