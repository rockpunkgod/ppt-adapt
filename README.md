# PPT Adapt

一个能够自主设计并制作整套可编辑PowerPoint的Agent Skill。根据主题、受众与内容选择艺术方向、字体、配色和逐页构图，再自然组合已有技术配方。

[下载最新源码与技能](https://github.com/rockpunkgod/ppt-adapt/archive/refs/heads/main.zip) · [下载可编辑 PPT](https://github.com/rockpunkgod/ppt-adapt/releases/download/v1.0.0/ppt-adapt-value.pptx) · [完整实播视频](https://github.com/rockpunkgod/ppt-adapt/releases/download/v1.0.0/ppt-adapt-playback.mp4) · [v1.0.0 发布归档](https://github.com/rockpunkgod/ppt-adapt/releases/tag/v1.0.0)

## 先看成品

**输入一段需求，交出三页连贯发布。**

> 为一款离线工作的边缘 AI 模块做三页产品发布，依次讲清：它是什么、怎样工作、用在哪里。

[![PPT Adapt 制作的三页产品发布：整体展示、工作方式与应用场景的原生连续播放](docs/media/ppt-adapt-demo.gif)](https://github.com/rockpunkgod/ppt-adapt/releases/download/v1.0.0/ppt-adapt-playback.mp4)

上方是本机 PowerPoint 实播节选，点击动图即可播放预览；[完整 55 秒实播视频](https://github.com/rockpunkgod/ppt-adapt/releases/download/v1.0.0/ppt-adapt-playback.mp4)可单独下载。文字、图表与页面结构保留可编辑对象；不是将整套页面做成视频。

<details>
<summary>展开查看完整 10 页</summary>

![PPT Adapt 项目介绍完整十页预览](examples/ppt-adapt-value/preview.jpg)

[作品说明](examples/ppt-adapt-value/README.md) · [原创主视觉与提示词](examples/ppt-adapt-value/assets/README.md)

</details>

## 设计方式

开放创意任务主动探索，艺术方向由内容与受众决定。随机联想可以打破模板惯性，最终用表达效果选择方向；品牌续稿和局部改稿保留必要约束。共用技术配方而不沿用固定美术模板，整套语言连贯，页面随叙事变化。见[制作总则](officecli-ppt-design/references/deck-design.md#开放创意任务主动探索)。

先围绕内容创造视觉主体、空间隐喻和观看顺序，再选配方实现：让设备与现场、材料与判断、整体与局部形成有意义的画面。通过位置、尺度、连接和连续镜头解释关系，让创意深入到主体与场景的构思。见[视觉构思方法](officecli-ppt-design/references/deck-design.md#从内容创造主体空间与镜头)。

开放需求先尝试两到三个不同方向的封面与代表内容页，依据真实预览比较；通过大胆的艺术方向和精准取舍完善作品，删去无用容器、过度强调和重复动画。

入口：[officecli-ppt-design](officecli-ppt-design/SKILL.md) · [整套演示设计指南](officecli-ppt-design/references/deck-design.md) · [配方索引](officecli-ppt-design/references/recipe-index.md)

正式版：[v1.0.0 发布与下载](https://github.com/rockpunkgod/ppt-adapt/releases/tag/v1.0.0) · [版本记录](CHANGELOG.md)

## 使用

**安装 Skill → 输入需求 → Agent 制作 → 收到 PPTX 和预览。**

让支持Skill安装的Agent安装本仓库中的`officecli-ppt-design`，或将该目录放入Agent的技能目录。安装后直接描述需求，例如：

> 用 PPT Adapt 为我的离线 AI 产品做一套9页发布演示，面向产品团队。下面是产品资料和图片。希望有鲜明的设计，讲清价值、工作方式和应用场景，给我可编辑PPTX和预览。

不需要指定OfficeCLI、COM、配方或脚本。Agent负责设计与制作，自动衔接内部工具并预览修正，再返回文件。可以指定品牌、场景、页数和参考风格，也可以交给Agent决定。

本地完整制作需要Windows桌面版PowerPoint及制作运行环境；Agent先检查并复用已安装的工具，只在实际缺失时说明需要准备什么。

### 一键重建首页作品

下载[最新仓库源码](https://github.com/rockpunkgod/ppt-adapt/archive/refs/heads/main.zip)，在具备运行环境的本机双击`rebuild-demo.cmd`，即可从随附源文件和本地素材重建十页介绍，输出**可编辑PPTX、整套预览图和动效预览视频**。不重新下载素材，不覆盖已有文件。该入口新增于v1.0.0之后，旧版压缩包不包含它。

也可以直接对Agent说：**“重建 PPT Adapt 的首页示例，把PPTX和预览给我。”** 技术参数由Agent处理，详见[内部重建说明](officecli-ppt-design/references/rebuild-demo.md)。该脚本只复现示例，新需求仍由Agent独立设计。

默认由同一Agent保持全稿设计上下文：理解需求 → 确定风格 → 规划整套演示 → 选择和组合配方 → 制作 → 实际预览 → 调整 → 交付。设计便笺按任务需要保留，不要求额外Agent、机器规格、依赖图、评分或兼容性数据库。

## 制作能力

- 内容决定艺术方向：商业证据、产品展示、人物叙事和文化提案采用适合自己的视觉语气。
- 字体角色、字距、网格、比例和色彩形成共同语言；页面按职责变化，保留非对称、大图、大字和章节停顿的自由。
- 将字体/网格、主体/遮罩/Morph、图表/强调/逐步揭示、3D场景、交互Zoom等已有机制组合成完整画面。
- 冲突在制作时通过重排、比例和时序处理；效果服务于内容，不凑技术数量。
- 先看关键画面，再看整套节奏，最后原生播放。特殊配方继续沿用其已有必要验证，不扩建审计平台。

保留OfficeCLI主体制作、PowerPoint COM/VBA精修及原生PowerPoint预览。既有配方与适用边界保持原样；仓库原有附加规划工具可按特定任务使用，不属于默认制作链。

## 原创作品

[PPT Adapt：从一段需求，到完整发布](examples/ppt-adapt-value/ppt-adapt-value.pptx)以一次实际任务介绍项目：从简短需求做出三页连贯的原创概念产品发布，再展示不同设计方向、可编辑图表与实际制作链。[作品与实播](examples/ppt-adapt-value/README.md)让项目价值落在具体成果上。

[PPT Adapt：每个想法，都有自己的样子](examples/ppt-adapt-unique/ppt-adapt-smooth.pptx)是新的10页独立设计版：松绿与信号黄、衬线排版、环形数据页及双面3D转化场景。先比较三个封面/内容页方向，再以减法精修收敛，并优化柔和过渡与入场节奏；[作品与预览](examples/ppt-adapt-unique/README.md)展示新的制作总则与实际验证范围。

[PPT Adapt：让 Agent 做出值得上台的 PPT](examples/ppt-adapt-showcase/ppt-adapt-showcase.pptx)是12页完整项目介绍，以演示本身体现设计与制作能力。包含内容组织、真实数据图表、连续3D聚光场景与可编辑交付；[原生播放预览](examples/ppt-adapt-showcase/playback.mp4)和[作品说明](examples/ppt-adapt-showcase/README.md)一并提供。

[循迹 TRACE：团队知识工作台](examples/trace-product/trace-product.pptx)是一套9页科技概念产品演示。采用编辑式字体层级、墨色/浅灰/酸橙色、多种构图、可编辑产品界面、原生图表及连续场景Morph。

[作品说明与制作源](examples/trace-product/README.md)。产品与数值均为概念演示，非在售功能或实测性能；跨机需要Noto Serif SC与Noto Sans SC，或替换后检查折行。

[动效组合PPTX](examples/motion-combinations/motion-combinations.pptx)展示斜切文字、同组多效果、固定导航/滚动时间轴/任务Zoom；[制作指导](officecli-ppt-design/references/advanced-composition.md)说明组合顺序、共享坐标、支点与适用边界。

[书页翻转PPTX](examples/book-flip/book-flip.pptx)展示曲线纸页围绕书脊的原生透视翻转，包含可编辑制作源及翻页、回退、再次翻页的本机播放记录。

[瓶形液面PPTX](examples/flask-wave/flask-wave.pptx)展示真孔遮罩、液面回弹及双层曲线波浪斜向上升，附可编辑制作源和本机播放记录。

[视频裁形PPTX](examples/video-intersect/video-intersect.pptx)提供裸视频与真媒体自定义几何对照，保持可编辑媒体及原生动态播放。

[原生GIF纹理PPTX](examples/native-gif-texture/gif-texture.pptx)展示透明活动纹理与真实裁形，另附可编辑动画源PPTX及原生播放。

[笔尖同步PPTX](examples/ink-pen-sync/ink-pen-sync.pptx)展示真实墨迹Replay与笔尖路径同步，以及原生入场、离场和返回后重新播放。

[逐字波浪PPTX](examples/text-wave/text-wave.pptx)保留单个可编辑文本框，以原生Wave展示字母错峰、旋转、幅度和连续循环。

[点击停转PPTX](examples/until-next-click/until-next-click.pptx)展示当前角度保持、后退重启，以及同样点击下持续旋转的对照。

[指定媒体触发PPTX](examples/media-shape-triggers/media-shape-triggers.pptx)将蛋糕/GIF与标题/视频分为独立入口，组合原生飞入、播放和70%透明，并提供不透明对照。

[原生3D动作PPTX](examples/native-3d-effects/native-3d-effects.pptx)组合真实3D入场、连续转盘及点击模型跳跃转身，附原创GLB及40秒原生播放记录。

[双首页导航PPTX](examples/double-home/double-home.pptx)让首次入场只播放一次，两个多页章节各完成两轮往返后，都回到静态首页，保留可编辑页面和原生播放记录。

[悬停章节导航PPTX](examples/hover-section-zoom/hover-section-zoom.pptx)将悬停预览接入真正多页Section Zoom，包含章节播放、节末返回、移出复位和重复进入，附连续原生录像。

[双向时间轴PPTX](examples/push-return/push-return.pptx)使用真实可编辑路线与原生Push，前进和返回时都保持跨页连接，附两轮方向键往返的连续原生录像。

[逐柱讲解PPTX](examples/chart-elements/chart-elements.pptx)在同一真实图表上比较系列内与类别内两种元素顺序，以自底部擦除逐根揭示数据，附连续原生录像。

[曲线关系图PPTX](examples/curve-links/curve-links.pptx)使用可编辑平滑路径、四段线渐变与原生投影表达类别间关系，节点与交叉层级清晰。

[椭圆顶盖字形PPTX](examples/ellipse-cap/ellipse-cap.pptx)将粗体字形映射到两端收尖的分割轮廓，保留字孔、可编辑路径和原始文字备份。

[玻璃霜边PPTX](examples/frost-glass/frost-glass.pptx)将原生局部模糊与白色20pt内阴影组合，保留清晰原图、原生模糊参考和可移动的背景填充窗口。

[字形局部虚化PPTX](examples/glyph-gaussian/glyph-gaussian.pptx)比较全字与单笔Gaussian，在深浅底色上保留清晰骨架；原文字和清晰路径可编辑，虚化图层可重建。

[灯具聚焦PPTX](examples/lamp-focus/lamp-focus.pptx)组合原创3D灯具、三段透明渐变光束、圆形图片与Morph，跨三页引导注意力；附原创资产和27秒原生前进、返回播放。

[烛焰呼吸PPTX](examples/flame-flicker/flame-flicker.pptx)用原生矢量与纵向70%错峰循环营造微动，文字保持静止；附静态对照与22秒本机播放。

[笔刷揭示PPTX](examples/brush-reveal/brush-reveal.pptx)从虚化底图逐笔擦出清晰内容，六笔交替方向后淡入标题；原生轮廓、背景填充和动画可编辑，附45秒实际播放。

宇航员3D场景的[制作指导](officecli-ppt-design/references/model3d.md)与[52秒原生播放](officecli-ppt-design/verification/native-3d-navigation/playback.mp4)也已入库：保留模型内部动作，组合跨页Morph及四圆点导航，8次非顺序点击通过。此项使用调用方提供的原生模板，技能包不附作者原模板或模型文件；区别于上方附原创GLB的3D动作示例。

[局部拖尾与冷暖叠加PPTX](examples/image-blast/image-blast.pptx)组合实际Blast透明条带与两个独立Overlay画笔层；文字和图片位置可编辑，附四页作品、原始图像层和27秒本机窗口放映。

## 环境

OfficeCLI、Python 3；本地预览与精修使用Windows桌面版PowerPoint和Windows PowerShell。按具体配方准备Pillow、NumPy或FFmpeg等依赖。字体与素材由调用方确认可用。

复用本机已经安装的工具、字体、Python环境和已有素材，不重复下载。Python依赖汇总见 `requirements.txt`；只在所选配方缺少依赖时补齐。几何配方沿用原有版本要求。OfficeCLI、FFmpeg、PowerPoint和字体属于外部运行环境，仓库不附带这些软件的安装文件；制作时向脚本传入本机已有工具路径。

## 完整技能包

仓库保留技能、配方、制作脚本、可编辑示例和已有原生播放记录。`ppt-skills.zip`保留为v1.0.0离线归档，也包含可选的`ppt-director`与`ppt-composition-planner`；新增重建入口请使用最新源码。默认制作入口仍为`officecli-ppt-design`。

维护者将新增项目文件加入Git后，可在本地执行 `python scripts/package_delivery.py` 更新离线包。打包只读取已跟踪的本地文件并检查压缩包完整性，不下载资源，不打入缓存、工具安装目录或工作区临时文件。归档的播放记录仅证明对应作品的既有结果，不能替代新作品或新组合的原生验证。

参考Taste Skill的主体优先、字体比例与克制，以及PPT Master的内容驱动规划、页面职责和整套视觉语言；没有移植网页CSS、响应式实现、SVG引擎或检查平台。
