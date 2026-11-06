# 已验证配方按实际后端分类

本表是方法索引，不是“每个来源全部效果通过”清单。一个视频可拆出多个配方；多个视频的同一机制合并复用。静态配方只验结构和画面，动效必须额外验中间态，交互必须实际点击。已验证指本轮列明边界内的样稿，不保证任何素材替换后自动通过。

多个配方一起使用时另读[高级动效组合](advanced-composition.md)：同组多动画、斜切字形/刀光、固定导航/滚动时间轴/任务Zoom、路径支点与3D翻页的制作顺序和边界。组合样稿在`examples/motion-combinations`，其通过范围单独列明，不改写下表原子配方的语义或验收标准。

`examples/book-flip`补充已实播的两页曲线纸页翻转：OfficeCLI基础+COM对称隐藏半页/Camera62/FOV60/Contrasting18/Morph2秒，15秒物理翻页/回退/再翻页。它是平面组的透视旋转组合，背面留白，不扩张为多页双面书本或GLB能力。

`examples/flask-wave`补充真瓶形Subtract/背景遮罩、Fly末段回弹与双曲线波浪斜向上升；18秒本机实播，覆盖边界与中间态已检查。详见[高级动效组合](advanced-composition.md)，不改变下表矩形窗口配方的范围。

## officecli-native

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|版式|网格、三级间距、视觉重量、主次变体|`build_layout_methods.py`：8页静态原生样稿，未扩张为海报课程所有外观|
|版式/叙事|对齐参照、留白、结论式图表、SCQA|`layout_story_build.py`：静态布局与叙事示例；内容结论仍须由真实材料支持|
|字体|标题/正文角色、结构和笔画对比配对|`type_build.py`：4页静态基础方法；缺字体时记录本机替代|
|字体|真文字多stop渐变含透明端|`text_gradient_*`：run gradFill5stop，双背景反解alpha/字孔、重开与改字48项，OfficeCLI原生；限已测单行Latin/Arial Black|
|中文/多图|轴线、留白、标题成块、图片窗口与主题环绕|`bili_layout_build.py`：13页样稿中的静态方法；瀑布流仅验证静态布局|
|表格|分组表头、横纵合并、必要边框和行列合计|`bili_layout_table_*`：真6×5表、横4/竖2合并、700合计、边框和原生尺寸/画面、包内前向过；不是自动Excel公式重算|
|图片|多个窗口继承同一父组图片坐标|`bili_layout_group_image_*`：真实父组blipFill/子grpFill，3页原生、8区域像素与包内前向过；限轴对齐静态矩形|
|图表|按系列/类别分步出现|`chart_build_*`：真实chart子目标，原CLI2/4次点击、6物理点击/8状态/12秒录屏过；保numLit数据，此入口不含xlsx/编辑数据|
|图表|目标/实际同轴100%重叠|[target-actual-overlay.md](target-actual-overlay.md)及`target_overlay_*`：真clustered双系列、零起点同主轴、目标100→120而实际柱保持；无xlsx能力声明|
|图表|逐根数据柱原生渐变|[chart-series-fill.md](chart-series-fill.md)及`chart_series_fill_build.py`：真实series gradFill、90度、逐柱归一，数据与0起点轴保持|
|3D/Morph|从零构造双面GLB卡片|[glb-double-sided-card.md](glb-double-sided-card.md)及`glb_card_*`：正反独立纹理/相反法线与绕序，真model3d，8秒实播两次物理点击与侧缘透视|
|媒体|保原始媒体的前后裁剪|[media-controls.md](media-controls.md)及`media_controls_*`：trimEnd为尾裁量，12秒保留2–10，实际PCM约8.05秒；不烘焙剪切|
|配色|选色系→定色值→调比例|`build_components.py --recipe palette`：自拟色值与面积示意，不规定60/30/10|
|配色/字体|主题色槽与major/minor EA/Latin引用|`theme_semantics_*`：两页三状态、146项/12原生PNG、保存重开保持；单改major EA不影响minor/Latin/direct，不代证复杂themeOverride|
|Morph|斜面照片、连续裁片、阅读视差|`build_morph.py`、`build_visual_animation.py`；提供自己的图片，检查裁切与正文清晰度|
|Morph|简化卡片抽拉、共支点花瓣|`build_components.py`的cards或petals配方：卡片数字叠放；花瓣为原生透镜路径，支点补偿已验|
|Morph|遮罩数字条、矩形焦点窗口|`build_components.py`的timeline或focus配方：固定窗口及边界无漏缝；不等同源圆孔聚光|
|Morph|真正空心圆聚焦|`build_circle_focus.py`：原生donut调孔，三状态平移，原生动态及选中/降暗像素核查；不是3D灯光场景|
|基础动画|反向飞入、逐项淡入、递增延迟|`build_visual_animation.py`及对应原生测试；实际触发顺序有独立检查|
|循环背景|双幅周期网格、同组线性滚动后复位|`seamless_grid_*`：17秒一页、真20%alpha、54秒1440实际帧/17项过，三窗速度误差<0.013%；限均匀周期网格|
|转场/时间轴|向上Push跨页路线接缝|[push-timeline-seam.md](push-timeline-seam.md)及`push_seam_*`：真开放路径同x贴边，3pt同色、两次上推，fresh真实片14中间态无断口；非反向返回能力|
|模板/时间轴|真master/layout/local独立动画|[master-layout-animation.md](master-layout-animation.md)及`master_animation_*`：三层独立timing、自动往返到页末、本地物理click不打断；18秒fresh及自然前进/显式Reset过|
|3D|已有原生模型与动画资产的保留改稿|`build_morph.py --mode 3d-template`：模板保留工作流；[四圆点导航](model3d.md)另经52秒/8次真实鼠标复测，不是从零创建库存模型|

增量原生配方：

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|图表|同值同心pie/ring与百分比卡片|[concentric-share-charts.md](concentric-share-charts.md)、`concentric_chart_*`：7真chart、CLI改值/标签同步、透背景与扇角；静态fresh过|
|交互/媒体|活动GIF与同源静态PNG配对|[gif-png-menu.md](gif-png-menu.md)、`gif_menu_*`：18结构/20实屏、六物理click与重入，真实多帧GIF；不承诺帧0重启|
|图片色彩|两端点原生Duotone|[native-duotone.md](native-duotone.md)、`duotone_*`：a:duotone、原图字节、12结构/10原生像素；彩色luma/透明合成另验|

原生仪表配方见[gauge-synchronized-values.md](gauge-synchronized-values.md)及`gauge_sync_*`：25/60驱动真实doughnut、共同支点Spin、两列滚字，11秒三次物理click与字形/角度/结束保持过；Timing2s与实屏1.5s未完/2.1s完成边界分别报告。

新增通过的原生矢量/合成配方：

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|字形轮廓|2/4/6毫米圆角等距扩张、4毫米描边转填充|[glyph-equal-offset.md](glyph-equal-offset.md)、`glyph_offset_*`；限定Arial Black 6/11厘米，逐点距离与5原生PNG过，非活文字|
|汉字设计|原创山川一笔共享、局部符号替换|[chinese-stroke-design.md](chinese-stroke-design.md)、`chinese_stroke_*`；7→6笔、独立几何与像素对照，非任意字体拆笔|
|笔画立体感|独立闭合补片、局部交叉换层、原生Multiply|[stroke-depth.md](stroke-depth.md)、`stroke_depth_*`；71项与3原生页/重开过，Bezier局部裁剪为明确细分近似|
|柔和色场|透明径向渐变与叠放顺序|[soft-local-color-field.md](soft-local-color-field.md)、`soft_field_*`；20椭圆240stop双次核对，颜色更改支持域之外像素不变|
|图片混合|原生Multiply/Overlay/Screen及局部Multiply|[blend-compositing.md](blend-compositing.md)、`blend_compositing_*`；只使用列明通过模式，Overlay50负例不纳入|
|扫描视觉去底|反相后固定底色Screen|[scanned-stroke-extraction.md](scanned-stroke-extraction.md)、`scan_stroke_screen_*`；3固定背景，输入仍不透明，非真alpha|

以上表归属`officecli-native`；COM仅承担原生验收。

## powerpoint-com

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|时间轴|按对象/效果序号改时长、延迟、加减速、反转与重复|`polish_timing.ps1`：已有MainSequence的精修，不能凭空建立指定对象触发|

## hybrid

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|字体|转曲、相减高光、条纹联合后相交阴影|`type_build.py`→`type_polish.ps1`：真实矢量布尔，保留转曲前基础稿|
|模板/编号|母版样式、真实动态页码字段、复制重排自动编号|[dynamic-slide-number.md](dynamic-slide-number.md)、`slide_number_field_*`；版式承接占位符、3→4页/重排/重开15原生PNG，非普通数字文本|
|字体|跨底色局部换色、真实数字剪角|`glyph_slices_*`：原位下半字形求交、右下角Subtract、独立斜线；包内2页前向、结构与原生像素过，静态结果不自动跟随背景|
|文字动画|同一框按自然段、英文字符或单词构建|`text_build_*`：三页，真paragraph范围、lt/wd iterate、Swivel200ms/10%；14秒实播502独立帧/34项前向过，不拆成多框|
|字体/表格|真字形描边、字形填图、1×5真表格交替透明度及顺序入场|`bili_layout_build.py`→`bili_layout_polish.ps1`：数字Float30、标题正文同一对象Fly2；源CLI误写Peek已修正。字形图UserPicture默认stretch须按素材另查比例|
|形状/Morph|凸出数字与卡片联合后组合平移|`native_card_build.py`→`native_card_polish.ps1`：真实Union，两页Morph；不是叠字近似|
|版式/阴影|柔边半椭圆、同色面板与折页分界|`shadow_fold_*`：真正柔边原生导PNG、裁半、等比、组合旋转；白/红两页静态与包内前向通过；不是均匀矩形阴影|
|形状/背景|字形使用全画幅幻灯片背景填充|`background_window_build.py`→`background_window_polish.ps1`：真实Fill.Background，静态像素坐标和移动中间态已验|
|形状/视频|真字形孔罩扩张穿越|[glyph-video-portal.md](glyph-video-portal.md)及`glyph_portal_*`：单字H真Subtract孔/90x Grow/前景退出/后置自动视频，11结构+5实屏项通过；多字逐字另验|
|形状/版式|真实布尔图片几何、原生占位框|`layout_story_geometry.py`及其精修：按该参考的具体通过项使用；剔除了有缺陷的Arc页|
|图表|同数据折线+渐变面积、原生数据工作簿|`chart_line_area_*`：同主轴同值，嵌xlsx与numRef/缓存一致；原生29→31编辑通过，默认前向保持29；立即恢复实验未通过，不声称恢复能力|
|图表|数据柱原生窄横纹填充|[chart-series-fill.md](chart-series-fill.md)及`chart_series_fill_polish.ps1`：真series pattFill/narHorz，前景绿/背景深蓝不透明，保存重开及像素条纹通过|
|图表|将共享类别完整numLit图表绑定真实xlsx|`chart_workbook.py`和`chart_workbook_polish.ps1`：单系列bar/双系列column默认视觉像素相同、实际改值副本及包内前向；不处理已有工作簿/散点/稀疏缓存|
|结构/图片|真SmartArt节点、父子层级、换布局、图文配对|`bili_layout_smartart_*`：四节点增删与Promote/Demote，保存重开及三布局；presAssocID→图片关系/hash独立核对，包内3页前向通过|
|矢量素材|简单双路径SVG转原生可编辑几何|[svg-native-shapes.md](svg-native-shapes.md)及`svg_native_*`：真custGeom/轮廓孔洞、单leaf改色、保存重开；原SVG保留，复杂SVG不外推|
|模板|真母版共享标识、CustomLayout分支与本地内容隔离|`bili_layout_master_*`：3页A/A/B、3状态9原生PNG、关系和像素差分及包内前向通过；不扩大为主题字体/母版动画能力|
|模板/文字|真Title/Body占位符继承、两级正文、局部覆盖与Reset|`placeholder_*`：四页六阶段24PNG、独立样式cascade/关系/文本/级别/像素隔离过；逐属性覆盖，换layout可重排几何|
|时间轴|7层时间拖尾、延迟0–420ms、往返与重复|`build_timeline.py`→`polish_timeline.ps1`：三轮原生播放和独立前向复测|
|时间轴/媒体|音频书签整组循环与尾段停顿|`audio_bookmark_*`：1交互组5效果，动作2s、音频2.5s；包内10秒真实录屏及像素相位通过；不含Ink Replay/跨页音乐|
|媒体|原生音频fade、视频bookmark定位|[media-controls.md](media-controls.md)及`media_controls_*`：同WAV同进程链fade0严格对照、原生Next/Next/Previous为6/9/6秒；不外推跨页生命周期或帧精确同步|
|媒体|跨页连续播放与有限停止|[media-cross-slide-lifecycle.md](media-cross-slide-lifecycle.md)及`media_lifecycle_*`：三页直线前进，第二页保持音阶连续、第三页停止，6实播+5结构项；不涵盖跳页/回退/无限loop|
|墨迹/媒体|真Ink逐笔Replay及5+2停顿循环|`ink_pause_*`：原创三trace/183个XYT、原生drawProgress5秒/7秒PCM书签；12项与29秒真实前向多轮重画/保持通过，不是Wipe图片|
|墨迹/路径组合|笔尖跟随Replay、Fly入场与离场|[笔尖同步](ink-pen-sync.md)及`ink_pen_*`：真Ink/透明圆心、5秒慢放及1秒串联，35秒实际点击与返回重播；限已测三等长trace|
|文字/循环|单文本框逐字Wave与幅度变化|[逐字Wave](text-wave.md)及`text_wave_*`：原生preset34/letter10%/一往返四旋转/父UntilEnd；44秒三参数与静态对照通过，整词周期含错峰|
|动画|连续自转、圆轨道、遮罩液面上升|`prepare_refinement.py`→`refine_visual.ps1`：圆轨径和往返原生检查；液面为矩形窗口，不是完整烧瓶波浪|
|动画|中心120%、两侧85%与50%透明保持；Wheel六节点相位|`wheel_focus_*`：父/子行为hold及7秒单辐条，经原生中间态与16秒真实录屏、修后包内前向8项独立审计含Fade时长|
|动画|原生Fly末段Bounce End|`bounce_end_*`：2秒Fly/Intensity0.75，GUI实读1.5秒；真实录屏238独立PTS帧，回弹与单调对照、13项审计通过；旧阅读器Fallback保留线性Fly|
|交互|透明实填充热点、内部导航、返回|时间轴/导航示例：实际点击通过，空白点击行为不扩大解释|
|交互|同页Pin首次开、二次关、第三次再开|`pin_toggle_*`：同一InteractiveSequence的Fade入口/出口、空白负对照、修后15项独立审计含500ms时长及真实渐变录屏；不声明任意快速连击|
|交互/教学|Kiosk答案反馈与Next依赖门控|`kiosk_quiz_*`：新题A67/B68项、各8秒实际放映，改题后旧正确不解锁/新正确解锁、Speaker负对照；非安全隔离|
|交互|鼠标悬停切换介绍状态|`hover_cards_*`：16条真hlinkHover、四卡五状态、15实播状态/12秒真实录屏/53项前向检查；空白保持，不含移出自动返回或点击深入|
|交互|悬停预览/孔内Zoom深入/移出返回|[hover-zoom-menu.md](hover-zoom-menu.md)及`hover_zoom_*`：5页真双孔/SlideZoom，两分支各重复两轮、连续真实MOVE与物理click，120秒/21状态/135项fresh过；不外推任意极速横扫|
|交互|Slide Zoom及返回首页|`repair_zoom.py`及对应原生检查：Choice和Fallback图像均补齐；不能称Section Zoom通过|
|交互|Section Zoom、透明背景、自动入口、节末返回|[section-zoom.md](section-zoom.md)及`section_zoom_*`：5页2节、4次真实点击、13秒桌面录屏、包内全新前向8项独立检查；旧软件Fallback仅静态单页链接|
|交互/目录|真Summary Zoom、原生Edit Summary与自动改列|[summary-zoom.md](summary-zoom.md)、`summary_zoom_*`：fresh CLI六页→原生UI真容器/2→3项，16秒/9物理click遍历三节并返回；需可见Normal等待缩略图，非手组SectionZoom|

新增hybrid整场景：

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|字形包络|自定义上下曲线带的受限包络映射|[bounded-envelope.md](bounded-envelope.md)、`envelope_band_*`；GDI原轮廓+CLI真实路径，17模型/6原生项，不支持任意拓扑或零高度尖端|
|轮廓混合|200中间级/30中间级、斜向挤出与连续长影|[contour-blend-extrusion.md](contour-blend-extrusion.md)、`contour_blend_*`；436面几何/颜色及COM真Union/Intersect重开过，不是可旋转3D模型|
|2.5D/媒体|真组支点、透视铰链、打开页视频层|[laptop-hinge-scene.md](laptop-hinge-scene.md)、`laptop_hinge_*`：22秒开合再开/8状态/60项，稳定屏内动态、屏外0差；过渡poster非实时3D视频贴面|
|图片/背景|原生艺术模糊与共享坐标局部窗|[local-blur-background.md](local-blur-background.md)、`local_blur_*`：原生Blur30缓存、清晰前景/真Fill.Background窗，12结构+12原生像素；严格Gaussian/动感另验|

视频自身裁形见[布局、叙事与几何](layout-story-geometry.md)末节：`examples/video-intersect`，OfficeCLI主体+COM Union/Intersect和播放恢复，24秒实播，真媒体custGeom、原视频字节及边界已验；不等同于后置视频孔罩。

原生透明GIF活动纹理见[native-gif-texture.md](native-gif-texture.md)：`examples/native-gif-texture`，可编辑动画源→真实1080p透明GIF导出→原生Intersect→等价fillRect修复，保存重开后18秒/多轮实播。此hybrid机制不扩大既有GIF/PNG菜单或指定触发配方。

## external-render

|设计用途|原子机制|入口与通过范围|
|---|---|---|
|限定混合公式|SoftLight、局部Overlay/SoftLight|[blend-compositing.md](blend-compositing.md)：外部编码RGB公式产生最终像素，PNG嵌入/原生合成已验；不宣称所有原生实现不可能|
|图像故障|真实R/GB通道分离与窄条复制错位|[image-channel-glitch.md](image-channel-glitch.md)、`image_glitch_*`；18离线/16原生项，明确失去通道效果的PPT内编辑性，Wind仅近似对照|
|限定纸墨提取|已知均匀纸色/单墨色输出真alpha|[scanned-stroke-extraction.md](scanned-stroke-extraction.md)、`scan_stroke_*`；原创模拟扫描与淡线/灰边过，不覆盖真实复杂纸纹阴影|

图片处理、CLI接口缺失和原生GUI能力不自动归到这一类；以上各项的核心像素机制实际发生在外部，保留原件、参数及生成脚本。

## 仍需验证的边界

任意新模板的3D菜单鼠标命中、3D照明场景/复杂纹理及其他仅有来源笔记的特殊组合，不能因基础机制相似而升级为已通过。新作品完成原生验证后再更新本索引。
