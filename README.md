# PPT 制作技能包

制作、优化和验证可编辑 PowerPoint 演示文稿。包含版式、字体、配色、图表、母版、Morph、路径动画、交互导航及 3D 场景配方。

整套演示按三个层次执行：

1. [ppt-director](ppt-director/SKILL.md)：先确定受众、叙事、全局视觉规范和整套页面规划。
2. [ppt-composition-planner](ppt-composition-planner/SKILL.md)：选择已有配方，编排对象、空间、时间轴和实现后端，处理冲突。
3. [officecli-ppt-design](officecli-ppt-design/SKILL.md)：制作可编辑页面，完成 PowerPoint 动效与原生播放验证。已有能力见[配方索引](officecli-ppt-design/references/recipe-index.md)。

## 使用

让支持 Skills 的 Agent 加载这三个技能，提供演示主题、内容、页数、受众及素材。例如：

> 使用 ppt-director 规划一份 8 页产品介绍，再用 ppt-composition-planner 组合已有配方，由 officecli-ppt-design 制作并用本机 PowerPoint 验证。采用简洁版式、清晰图表和适量 Morph，保留可编辑文字。

将三个同名文件夹添加到技能库，保持各自内部资源的相对位置。总控和编排层复用已有执行配方，不包含另一套演示引擎。

冲突按内容正确与完整性、可读性、整套视觉一致性、动效质量、装饰效果的顺序解决。章节可有明确设计的变化，整套一致不要求所有页面布局相同。单项配方通过不代表组合通过；未经整套原生实测的组合保持未验证。

## 实现方式

| 类型 | 用途 |
| --- | --- |
| `officecli-native` | 页面、排版、形状、图表、基础动画及 Morph |
| `powerpoint-com` | 原生动画的时长、延迟、加减速和行为精修 |
| `hybrid` | OfficeCLI 搭建基础，再用 PowerPoint COM 或必要的原生界面完成效果 |
| `external-render` | 核心图像或视频效果由外部计算或渲染生成，再嵌入 PPT |

## 环境

- OfficeCLI；本地验证使用 Windows 桌面版 Microsoft PowerPoint。
- Python 3 与 Windows PowerShell；按配方准备 NumPy、Pillow、SciPy 等依赖。
- 录屏和媒体分析需要 FFmpeg/FFprobe；字形几何配方另需 `scripts/requirements-geometry.txt`。
- 图片、字体和模板由调用方提供并确认可用。不同配方的输入限制与验证范围见对应说明。

技能目录保留可复用脚本、必要素材及使用说明。`examples` 保存原创集成验收的可编辑作品、组合计划、冲突记录和原生验证结果；研究缓存和失败试验不作为交付。未通过的效果不作为已验证配方。

[8页原创集成案例](examples/library-night-pilot/README.md)包含字体与主题、原生图表、原生表格和Morph组合；[本次原生验收](examples/library-night-pilot/integration-acceptance.json)339项通过，[组合状态](examples/library-night-pilot/composition-verified/composition-plan.json)已绑定本次成品与证据。该通过仅覆盖这个案例及其明确复用范围。
