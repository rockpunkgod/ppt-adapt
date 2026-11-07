# PPT Adapt

让 Agent 从你的内容出发，自主设计并制作整套可编辑 PowerPoint。

[下载技能与源码](https://github.com/rockpunkgod/ppt-adapt/archive/refs/heads/main.zip) · [下载当前 PPTX](https://github.com/rockpunkgod/ppt-adapt/raw/refs/heads/main/examples/ppt-adapt/ppt-adapt.pptx) · [观看当前实播](examples/ppt-adapt/playback.mp4)

## 先看成品

**让想法，抵达观众。**

从一句城市步行工具的需求，进入原创“漫游”提案：沿可编辑街区从全景走向路线，再靠近书店细节。章节、景别和注意力形成连续的观看路线；再以已完成的 RoboMaster 作品展示不同内容的设计差异，最终落到可以继续修改的 PPTX。

[![当前作品的原生连续播放节选](docs/media/ppt-adapt-demo.gif)](examples/ppt-adapt/playback.mp4)

点击动图播放节选；[完整原生实播](examples/ppt-adapt/playback.mp4)与[可编辑 PPTX](examples/ppt-adapt/ppt-adapt.pptx)可单独查看。

<details>
<summary>展开查看完整九页</summary>

![当前九页演示预览](examples/ppt-adapt/preview.jpg)

[当前作品说明](examples/ppt-adapt/README.md)

</details>

首页始终只展示当前作品；更新成稿时同步替换PPTX、预览、实播与重建源。

## 使用

**安装 Skill → 输入需求 → Agent 制作 → 收到 PPTX 和预览。**

让支持Skill的Agent安装本仓库的`officecli-ppt-design`，或将该目录放入Agent的技能目录。然后直接描述需求：

> 用 PPT Adapt 介绍我的产品，面向产品团队。这里是产品资料和图片。希望有独特的设计，把价值、工作方式和应用场景讲清楚，给我可编辑PPTX和预览。

可以指定品牌、页数、场景和参考方向，也可以交给Agent决定。Agent负责艺术方向、叙事、主体构思、页面制作与预览修正，内部处理工具衔接。

完整本地制作需要Windows桌面版PowerPoint与制作运行环境。Agent先检查并复用已有工具、字体和素材，只在实际缺失时说明需要准备什么。

## 一键重建当前作品

下载源码后，在具备运行环境的本机双击[`rebuild-demo.cmd`](rebuild-demo.cmd)，从本地制作源重新生成当前九页演示，输出可编辑PPTX、整套预览图和原生导出视频。不重新下载素材，不覆盖已有输出目录。

也可以对Agent说：**“重建 PPT Adapt 的当前示例，把PPTX和预览给我。”** 参数由Agent处理。脚本仅复现当前示例，新需求仍独立设计。内部使用方式见[重建说明](officecli-ppt-design/references/rebuild-demo.md)。

## 设计与制作

开放创意任务主动探索，艺术方向由内容与受众决定。先围绕内容创造视觉主体、空间隐喻和观看顺序，再选择已有机制实现。随机联想用于打破惯性，品牌续稿和局部改稿保留必要约束。

先写简短视觉分镜，规划章节场景、景别、焦点交接与关键高潮，优先试做最难的连续片段，确定方向后扩展整套。全稿保持共同的字体、色彩与空间语言，页面随信息关系变化；通过实际预览与减法精修完善作品。

所有素材先筛选、看原素材，再在实际版式里预览。来源可靠之外，还要确认主体清楚、内容准确、画质够用、风格合适、裁切后仍能表达；不以官网来源、文件尺寸或生成成功代替选择。

主体制作沿用OfficeCLI，必要时由PowerPoint COM/VBA精修，最后用原生PowerPoint预览和播放。已有字体、版式、图表、Morph、3D、媒体和交互配方继续复用，不新建PPT引擎或审美评分平台。

[技能入口](officecli-ppt-design/SKILL.md) · [设计总则与视觉构思](officecli-ppt-design/references/deck-design.md) · [内部配方索引](officecli-ppt-design/references/recipe-index.md)

运行依赖按所选配方准备，见`requirements.txt`与内部参考；OfficeCLI、PowerPoint和字体属于外部运行环境。复用本机素材和工具，不重复下载。历史发布与改动保留在[版本记录](CHANGELOG.md)，首页不陈列迭代作品。
