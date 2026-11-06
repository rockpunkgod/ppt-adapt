---
name: officecli-ppt-design
description: 自主设计、制作和优化整套可编辑PPT。先理解受众与内容，确定艺术方向、字体配色和逐页节奏，再组合已有版式、图表、Morph、动画及3D配方；以OfficeCLI制作、PowerPoint COM/VBA精修和原生播放预览完成交付。
---

# PPT Adapt：整套演示设计与制作

工作链：**理解需求与内容 → 确定艺术方向 → 规划整套演示 → 组合已有配方 → OfficeCLI制作 → 实际预览与调整 → PowerPoint原生播放 → 交付。**

本技能补充OfficeCLI，主指令不被替换。依据调用方提供的内容制作；示例占位文字、素材与题材须替换成实际需求。

整套制作先读[整套演示设计指南](references/deck-design.md)。由当前Agent保持全稿上下文，自然完成内容、风格、页面和动效选择；常规制作无需额外Agent、JSON规格、依赖图或配方兼容性数据库。设计便笺只记值得保持的决定，复杂程度随任务增加。

开工前想清楚：观众要理解或决定什么；用什么视觉方向表达；哪些字体、颜色、边线和元素贯穿全稿；每页的一句话与主体画面；哪些相邻页需要连续、哪些需要停顿。先通读全套论点，再进入逐页制作。品牌、事实和用户约束优先；有意的章节变化保留理由，不把统一做成同一版式的重复。

组合时先定页面骨架，再加有意义的视觉处理，最后安排连续对象与时序。字体层级、网格、大图、遮罩、图表、3D和Morph都从现有参考中选；一个效果负责一个表达任务，达到目的就停。预览封面、主要证据页和关键连续场景，看到比例、字距、换行或运动不对就直接改作品。

涉及同对象多动画、斜切文字、滚动时间轴与Zoom、路径支点、3D翻页或瓶形波浪组合时，读[高级动效组合](references/advanced-composition.md)。先确定共享对象、固定/运动层和点击组，再追加效果；组合的实播结论独立于所复用的单项配方。

## 按实现后端封装

每个原子配方使用用户指定的四类之一：`officecli-native`（CLI可完成）、`powerpoint-com`（既有文稿的原生精修）、`hybrid`（CLI基础+COM几何/动效）、`external-render`（核心视觉由外部图像计算或AE等渲染产生）。判据与交付要求见 [execution-backends.md](references/execution-backends.md)。设计方法类别作为第二维标签，不能把CLI缺口直接判为必须外部渲染。

可用方法按后端与设计用途列在[配方分类索引](references/recipe-index.md)。先匹配具体机制，再加载相应参考，避免为相同能力建立重复流程。

## 先理解，再选方法

从现有内容确定页面的结论、受众、信息关系与交付环境。并列信息适合分组或卡片；时间变化适合连续轨道；对比聚焦适合阅读面板；长篇正文优先清晰静态结构。不为“高级感”自动添加所有动画。

有参考作品时，先确认其构图、对象关系、运动轨迹和节奏；单张截图不能证明中间动画。缺少内容或素材时列明所需输入，不虚构数据、素材授权或已验证的效果。

## 按实际机制选择参考

按所需版式、对象关系和动效机制选择配方，保留每项的适用范围。

|需要|阅读|可复现起点|
|---|---|---|
|构图、图像窗口、阅读面板、配色|[构图与配色](references/layout-color.md)、[配色方法](references/color-method.md)|`build_visual_animation.py`的连续裁片/文本列|
|网格、间距、视觉重量、同文案不同主次|[网格与间距](references/grid-spacing-methods.md)|`build_layout_methods.py`，8页可编辑代表方法|
|对齐参照、SCQA叙事、原生图表、布尔图片几何|[布局、叙事与几何](references/layout-story-geometry.md)|`layout_story_build.py`、`layout_story_geometry.py`与精修脚本|
|同一度量的折线+渐变面积、原生编辑数据|[原生组合图](references/chart-line-area.md)|`chart_line_area_*`：同轴同值、嵌入工作簿及实际改值测试；默认不修改示范数据|
|已有字面量bar/column图补真实“编辑数据”工作簿|[图表工作簿补齐](references/chart-workbook-materialization.md)|`chart_workbook.py`与精修脚本：当前限每文稿一图，默认保数据和画面；不清理已有工作簿|
|真实图表按系列或类别分步讲解|[原生图表构建](references/chart-build.md)|`chart_build_*`：OfficeCLI生成真实chart子目标，series两步/category四步；此配方不包含“编辑数据”工作簿|
|同轴比较目标与实际|[目标实际重叠柱](references/target-actual-overlay.md)|`target_overlay_*`：真clustered chart、100%重叠、透明目标在后；改目标而实际高度保持|
|真实数据柱内部渐变与横纹|[图表系列填充](references/chart-series-fill.md)|`chart_series_fill_*`：CLI逐柱渐变、COM原生pattern，保数据与零轴；保存重开像素一致|
|真正SmartArt层级、增删节点、换布局及图文配对|[原生SmartArt](references/smartart-native.md)|`bili_layout_smartart_*`：保留diagram数据模型；三个布局实际测试，不用普通group代替|
|原生数据表、分组表头、真实合并和必要边框|[原生数据表](references/native-data-table.md)|`bili_layout_table_*`：OfficeCLI直接制作6×5真表与行列合计，COM仅只读核验|
|多片窗口共用完整图片坐标|[父组图片填充](references/group-image-fill.md)|`bili_layout_group_image_*`：真实父组blipFill与子形状grpFill；轴对齐窗口静态验证|
|调用方简单SVG转真可编辑路径|[SVG原生形状转换](references/svg-native-shapes.md)|`svg_native_*`：原生SVGEdit、两路径真custGeom、孔洞透底、单路径改色及重开；复杂SVG另验|
|共同母版标识与A/B角色版式继承|[母版与CustomLayout](references/master-layout-inheritance.md)|`bili_layout_master_*`：真实绑定、共享更新和兄弟layout隔离；不以三份复制形状代替继承|
|真标题/正文占位符、层级样式与本地覆盖|[占位符继承与Reset](references/placeholder-inheritance.md)|`placeholder_*`：真Title/Body、两级正文、换layout/原生Reset，六阶段24原生PNG验证|
|母版、布局、本地各自独立动画|[继承动画时间轴](references/master-layout-animation.md)|`master_animation_*`：真实master/layout循环、本地点击，18秒实播与跨页/显式Reset验证|
|主题色与中英标题/正文字体统一修改|[主题语义与直接格式隔离](references/theme-semantics.md)|`theme_semantics_*`：schemeClr与major/minor EA/Latin引用，局部direct格式保持；OfficeCLI原生制作|
|字体角色、形态搭配、白边高光与条纹矢量文字|[字体与基础设计](references/typography-foundations.md)|`type_build.py`与`type_polish.ps1`，保留转曲前文字|
|保留可编辑文字的多色标渐变和透明端|[文字渐变](references/text-gradient.md)|`text_gradient_*`：真run gradFill、双背景/字孔/重开/真实改字48项；不是整框渐变|
|跨底色局部换色、数字真实剪角|[局部字形与剪角](references/glyph-slices.md)|`glyph_slices_*`：原位副本求交、真实Subtract和独立细斜线；静态矢量，不自动随背景换色|
|中文文字、字体字面、多图与透明表格顺序入场|[中文与多图版式](references/bilibili-layout-type.md)|`bili_layout_build.py`与`bili_layout_polish.ps1`，真表格/字形描边/Float修正|
|单文本框逐段、逐字符、逐单词出现|[原生文本构建](references/text-unit-build.md)|`text_build_*`：真paragraph范围与Swivel文本iterate，保留原文字，不拆成多个框|
|单文本框逐字起伏、旋转与连续循环|[原生逐字Wave](references/text-wave.md)|`text_wave_*`：保留活文字、letter10%与四段旋转，0.5秒/2秒及两幅度44秒实播|
|照片揭示、斜切信息层、视差、时间轴|[连续状态与Morph](references/morph-patterns.md)|`build_morph.py`、`build_timeline.py`|
|卡片抽拉、背景字形窗口、花瓣展开、聚焦、滚动数字|[原生组件配方](references/native-components.md)|`build_components.py`；真实数字Union用`native_card_*`，全画幅背景用`background_window_*`|
|柔边单侧阴影、折页数字分界|[柔边折页](references/shadow-fold.md)|`shadow_fold_*`：原生导出透明椭圆后裁半，邻接同底色面板组合旋转|
|逐项点击、递增延迟、往返、环轨、拖尾|[时间与路径](references/timing-patterns.md)|`build_visual_animation.py`、`build_timeline.py`|
|跨页上推后时间轴路线连续|[Push接缝](references/push-timeline-seam.md)|`push_seam_*`：真开放路径同x贴边、两次原生Push，14中间态接缝量测|
|连续向左的周期网格背景|[双幅网格循环](references/seamless-grid-loop.md)|`seamless_grid_*`：双幅同组移一页、17秒无限repeat；54秒真实录屏与相位连续性验证|
|中心强调、两侧缩小降权、圆环逐项绘出|[聚焦与Wheel相位](references/wheel-focus.md)|`wheel_focus_*`：真实Grow/Shrink、透明保持、7秒单辐条与六节点相位|
|飞入末段回弹及衰减收敛|[原生Bounce End](references/bounce-end.md)|`bounce_end_*`：原生Fly的末段精修，GUI读值、真实录屏与无回弹对照轨迹验证|
|再次单击停在当前角度、后退重启|[点击相位保持](references/until-next-click.md)|`until_next_*`：原生Spin/onNext停止/Fade，32秒两轮停转与15秒UntilEnd对照实播|
|同页指定按钮开／关标签|[Pin指定触发](references/pin-toggle.md)|`pin_toggle_*`：首次开、二次关、三次再开，空白负对照；不是普通Next|
|答对后才显示继续按钮、改题迁移依赖|[Kiosk问答门控](references/kiosk-quiz.md)|`kiosk_quiz_*`：三答案指定触发、真实Next迁组与两题135项；Kiosk阻本配方的空白/方向键绕过|
|仅移动鼠标切换介绍卡片|[悬停状态导航](references/hover-state-navigation.md)|`hover_cards_*`：真实MouseOver内部动作，四卡五状态，移出保持；不是点击触发|
|悬停预览、移出返回、点击Zoom深入|[完整悬停Zoom菜单](references/hover-zoom-menu.md)|`hover_zoom_*`：真孔罩/SlideZoom、连续真实MOVE，120秒/21状态/135项fresh验证|
|多个不同时长的动作整组循环并留停顿|[音频书签序列](references/audio-bookmark-sequence.md)|`audio_bookmark_*`：一条书签触发首动作，同组WithPrevious推进，实际多轮录屏；不等同Ink Replay|
|同页指定蛋糕显示GIF、指定标题启动视频叠加|[媒体指定触发](references/media-shape-triggers.md)|`media_trigger_*`：两独立入口，Fly/Play/透明70%同组，36+20秒原生实播与不透明对照|
|媒体裁剪、真实声音淡化与视频书签跳点|[原生媒体控制](references/media-controls.md)|`media_controls_*`：保原媒体、同WAV对照实录、6/9/6跳点；跨页音乐另验|
|跨页连续音乐与有限停止|[跨页音频生命周期](references/media-cross-slide-lifecycle.md)|`media_lifecycle_*`：第一页启动、第二页不断不重启、第三页停止；仅已测直线翻页路径|
|真实墨迹逐笔回放、5秒绘制与2秒停顿循环|[Ink Replay停顿循环](references/ink-replay-pause.md)|`ink_pause_*`：原创InkML、保留原生drawProgress行为、7秒媒体书签；29秒真实多轮前向通过|
|笔尖跟随真墨迹，串联入场与离场|[笔尖同步](references/ink-pen-sync.md)|`ink_pen_*`：复用Replay，透明圆定位中心，同步路径，快放与返回后重播已实测|
|真字形孔罩扩张、文字退出与视频揭示|[字形视频穿越](references/glyph-video-portal.md)|`glyph_portal_*`：真实Subtract孔、90倍原生放大、后置自动视频；单字H前向实播通过|
|内部链接、点击区域、Zoom与返回|[交互导航](references/interaction-patterns.md)、[原生Section Zoom](references/section-zoom.md)|`section_zoom_*`：真实分节进入、自动节首、透明背景与节末返回；不以单页Zoom代替|
|原生动画导透明GIF、裁形后循环纹理|[原生GIF纹理](references/native-gif-texture.md)|`native_gif_*`：真Export与Intersect，fillRect恢复动态且保持画面比例|
|视频自身裁为自定义轮廓并保持动态|[视频原生相交](references/layout-story-geometry.md)|`examples/video-intersect`，真媒体custGeom；相交后恢复唯一Play，保存重开实播|
|真实3D模型、内部动作、跨页姿态|[3D场景](references/model3d.md)|`build_morph.py --mode 3d-template`，使用合法取得的模板|
|正反面不同内容的真3D卡片|[双面GLB卡片](references/glb-double-sided-card.md)|`glb_card_*`：自建双单面贴图、OfficeCLI真model3d/Morph、两次真实点击及透视侧缘验证|

只将完成适用验证的实现当作可直接复用配方；未实播的新组合不作为已有能力。具体限制见各配方。

同心饼图/外环与百分比卡片见[份额图表一致性](references/concentric-share-charts.md)：`concentric_chart_*`核7真chart、同步改值、真透缝和余数。活动/静态菜单见[GIF/PNG状态配对](references/gif-png-menu.md)：`gif_menu_*`保真实多帧GIF，以六次物理点击验重复进入。图片双色映射见[原生双色调](references/native-duotone.md)：`duotone_*`保原图字节，原生a:duotone端点与中间灰阶通过；不外推彩色亮度权重或精确GIF重启帧。

仪表联动见[半环、指针与滚动数值](references/gauge-synchronized-values.md)：`gauge_sync_*`核同一参数终态与动作后保持，滚字是装饰过渡。完整开合场景见[笔记本铰链与视频](references/laptop-hinge-scene.md)：`laptop_hinge_*`在过渡中使用静帧，稳定打开后叠原生视频。局部虚实见[原生模糊背景窗口](references/local-blur-background.md)：`local_blur_*`用原生艺术Blur30产生缓存，再以同世界坐标窗口展示；不是实时背景模糊参数或严格Gaussian保证。

## 两层制作

### OfficeCLI：基础制作

执行 `officecli --version` 并实际创建一个临时小样确认运行环境。静态页读本机 `officecli load_skill pptx`，连续状态读 `officecli load_skill morph-ppt`。对不确定属性先查本机help；帮助描述与实际写入不一致时以成品和PowerPoint行为为准。

本次Windows机器实测程序为`E:/ppt skills/tools/officecli.exe`（1.0.160）。存在时可显式传入`--officecli`；PATH中的另一个1.0.159曾在实际create时报缺少System.Private.Xml，版本号能显示不代表可用。其他机器重新发现程序，不依赖此本机路径，也不擅自改全局PATH。

- 页面、文字、形状、图片、排版、组合、基础动画和Morph由OfficeCLI制作；Python只用于坐标计算、CLI编排、素材处理和检查，不改用其他PPT生成库。
- 大多数效果先建立完整终态，再克隆起态/中间态；同一对象相邻页保持唯一的`!!name`。同页同名冲突应检查有效MC分支，而非误判Choice/Fallback副本。
- 文本和背景分层；面板可以变宽，正文重新排版，不能横向拉伸字。图片按本征比例裁切，不压扁来凑框。
- 以带单位坐标和UTF-8 batch JSON构建，检查每项失败与回滚。`add`返回的稳定ID用于后续编辑；组合后重新读取路径。
- 在非OfficeCLI程序读取文件前执行save/close。默认输出新文件，不覆盖原稿。

具体对象/命令映射与实测陷阱见 [officecli-bridge.md](references/officecli-bridge.md)。

### PowerPoint COM/VBA：动效精修

阅读 [powerpoint-polish.md](references/powerpoint-polish.md)。COM/VBA是正式精修层，用于精细Timing、加减速、自动反转、重复、指定对象触发和实际暴露的特殊动画，不局限于验收。

保留基础PPT、精修后PPT、参数计划与脚本；按对象名称和Effect定位，不重建或覆盖整页时间轴。CLI时长是毫秒，COM Timing是秒。相位、重复和触发须在PowerPoint里播放，不能用字段成功回读代替。

`polish_timing.ps1`可精修现有MainSequence效果；复杂触发器使用经验证的专用脚本。它不创建不存在的效果，也不把onClick变成指定对象触发。

## 可运行示例

新增静态机制按目标进入：[字形等距扩张](references/glyph-equal-offset.md)、[汉字共享笔画](references/chinese-stroke-design.md)、[逐笔画局部层次](references/stroke-depth.md)、[受限曲线包络](references/bounded-envelope.md)、[轮廓混合与长影](references/contour-blend-extrusion.md)、[柔和局部色场](references/soft-local-color-field.md)。这些配方保留明确的字形/路径/拓扑边界，不据此宣称任意字体自动设计。

需要复制/重排后自动编号时使用[真实动态页码](references/dynamic-slide-number.md)：母版样式、版式页码占位符和原生slidenum字段共同工作，不能换成固定数字文本。

自动生成章节目录、整体改列与节末返回使用[真Summary Zoom](references/summary-zoom.md)：OfficeCLI基础页加本机原生UI创建/编辑摘要容器；不要将手工SectionZoom组称为Summary，新增缩略图须等待可见Normal原生渲染就绪。

图片核心机制分开路由：[混合模式](references/blend-compositing.md)区分原生Multiply/Overlay/Screen与外部SoftLight；[通道分离/条带故障](references/image-channel-glitch.md)为外部像素处理；[扫描笔触](references/scanned-stroke-extraction.md)区分原生固定底色Screen与限定纸墨的外部真alpha。不得将视觉去底称作透明图片。

脚本不捆绑作者的模板、GLB或照片。现代/电影感示例要显式提供图片；3D示例是保留已有原生动画资产的模板改造，不冒称从零制作库存宇航员。

```text
python scripts/build_morph.py --mode modern --officecli officecli --outdir ./output --landscape /path/wide.jpg --portrait /path/portrait.jpg
python scripts/build_morph.py --mode cinematic --officecli officecli --outdir ./output --landscape /path/wide.jpg
python scripts/build_visual_animation.py --officecli officecli --outdir ./output
python scripts/build_timeline.py --officecli officecli --out ./output/timeline.pptx
python scripts/build_components.py --recipe cards --officecli officecli --out ./output/cards.pptx
powershell -NoProfile -File scripts/polish_timeline.ps1 -Deck ./output/timeline.pptx -Output ./output/timeline-polished.pptx
python scripts/verify_timeline.py ./output/timeline-polished.pptx
```

先运行所需脚本`--help`。路径中的`python`指可用Python 3；PowerShell脚本使用本机Windows PowerShell与PowerPoint。生成参数、素材和配方可随任务修改，修改后的部分需要重新验证。

## 本机验证与入库门槛

阅读 [validation.md](references/validation.md)。

1. **结构：** OfficeCLI保存后验证schema、关系、名称、转场与动画节点；审阅issues并区分有意离场和实际溢出。
2. **静态：** 本机PowerPoint原生打开、导出所有状态PNG，检查字体、裁切、留白、图像比例、文字对比和遮挡。
3. **动态：** 原生放映并录屏/导出视频，比较源目标的轨迹、节奏、连续性和视觉层级；查看转场前中后。循环测试要覆盖完整声明周期。
4. **交互：** 点实际目标，再测试返回、重复进入和空白点击。确认看到的是真正投影窗口，不能把Presenter View的错误坐标点击当功能失败。
5. **修正：** 将实际失败和修复写回配方，复测受影响部分。命令成功、schema通过、静态截图清晰均不能单独算动效通过。
6. **入库：** 只有完成上述适用测试的能力进入可用配方；研究观察和待验功能保留在独立待办。报告通过的是完整复刻还是明确简化的机制。

当前Agent保持整套演示的设计上下文；同一台PowerPoint的放映/导出串行，只关闭本次presentation，不退出共享PowerPoint或关闭用户文稿。

## 交付

默认交付可编辑PPT；需要继续修改时附基础与精修脚本。按调用方要求保留文件，任务结束清理缓存、试验稿和临时录屏。未通过的效果不冒充可用能力，不将实现细节塞入演示正文。
