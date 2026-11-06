# PPT Adapt

一个能够自主设计并制作整套可编辑PowerPoint的Agent Skill。根据主题、受众与内容选择艺术方向、字体、配色和逐页构图，再自然组合已有技术配方。

入口：[officecli-ppt-design](officecli-ppt-design/SKILL.md) · [整套演示设计指南](officecli-ppt-design/references/deck-design.md) · [配方索引](officecli-ppt-design/references/recipe-index.md)

## 使用

加载 `officecli-ppt-design`，提供内容、受众、场景、页数以及可用素材。例如：

> 制作一套9页科技产品介绍，面向产品团队。先确定整套艺术方向和叙事节奏，组合已有字体、网格、原生图表及适量Morph，保留可编辑对象，制作中预览关键页面，最后用本机PowerPoint播放检查。

默认由同一Agent保持全稿设计上下文：理解需求 → 确定风格 → 规划整套演示 → 选择和组合配方 → 制作 → 实际预览 → 调整 → 交付。设计便笺按任务需要保留，不要求额外Agent、机器规格、依赖图、评分或兼容性数据库。

## 制作能力

- 内容决定艺术方向：商业证据、产品展示、人物叙事和文化提案采用适合自己的视觉语气。
- 字体角色、字距、网格、比例和色彩形成共同语言；页面按职责变化，保留非对称、大图、大字和章节停顿的自由。
- 将字体/网格、主体/遮罩/Morph、图表/强调/逐步揭示、3D场景、交互Zoom等已有机制组合成完整画面。
- 冲突在制作时通过重排、比例和时序处理；效果服务于内容，不凑技术数量。
- 先看关键画面，再看整套节奏，最后原生播放。特殊配方继续沿用其已有必要验证，不扩建审计平台。

保留OfficeCLI主体制作、PowerPoint COM/VBA精修及原生PowerPoint预览。既有配方与适用边界保持原样；仓库原有附加规划工具可按特定任务使用，不属于默认制作链。

## 原创作品

[循迹 TRACE：团队知识工作台](examples/trace-product/trace-product.pptx)是一套9页科技概念产品演示。采用编辑式字体层级、墨色/浅灰/酸橙色、多种构图、可编辑产品界面、原生图表及连续场景Morph。

[作品说明与制作源](examples/trace-product/README.md)。产品与数值均为概念演示，非在售功能或实测性能；跨机需要Noto Serif SC与Noto Sans SC，或替换后检查折行。

[动效组合PPTX](examples/motion-combinations/motion-combinations.pptx)展示斜切文字、同组多效果、固定导航/滚动时间轴/任务Zoom；[制作指导](officecli-ppt-design/references/advanced-composition.md)说明组合顺序、共享坐标、支点与适用边界。

[书页翻转PPTX](examples/book-flip/book-flip.pptx)展示曲线纸页围绕书脊的原生透视翻转，包含可编辑制作源及翻页、回退、再次翻页的本机播放记录。

[瓶形液面PPTX](examples/flask-wave/flask-wave.pptx)展示真孔遮罩、液面回弹及双层曲线波浪斜向上升，附可编辑制作源和本机播放记录。

[视频裁形PPTX](examples/video-intersect/video-intersect.pptx)提供裸视频与真媒体自定义几何对照，保持可编辑媒体及原生动态播放。

[原生GIF纹理PPTX](examples/native-gif-texture/gif-texture.pptx)展示透明活动纹理与真实裁形，另附可编辑动画源PPTX及原生播放。

[笔尖同步PPTX](examples/ink-pen-sync/ink-pen-sync.pptx)展示真实墨迹Replay与笔尖路径同步，以及原生入场、离场和返回后重新播放。

[逐字波浪PPTX](examples/text-wave/text-wave.pptx)保留单个可编辑文本框，以原生Wave展示字母错峰、旋转、幅度和连续循环。

## 环境

OfficeCLI、Python 3；本地预览与精修使用Windows桌面版PowerPoint和Windows PowerShell。按具体配方准备Pillow、NumPy或FFmpeg等依赖。字体与素材由调用方确认可用。

参考Taste Skill的主体优先、字体比例与克制，以及PPT Master的内容驱动规划、页面职责和整套视觉语言；没有移植网页CSS、响应式实现、SVG引擎或检查平台。
