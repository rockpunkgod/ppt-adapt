# 布局、叙事与布尔图片几何：已通过配方

五来源完整记录在工作区distill/youtube-layout-story/source-notes.md；共1876段字幕全文与91操作帧已审阅。2小时长课字幕从0:00到2:01:19，片长2:01:30且有结束语；不称正常速度逐秒看完。下列保留通过子集，待验效果不升级为能力。

基础builder的图表使用numLit/strLit，保留原生chart和CLI可修改数值；交付PowerPoint“编辑数据”时，再按[真实工作簿补齐](chart-workbook-materialization.md)运行COM绑定。visual-design与storytelling两处图表现已分别通过真实xlsx/公式/cache一致、原生改值及默认视觉保留；默认保持原数值，改值测试另存副本，不由一次编辑推断立即恢复能力。不能凭chart节点存在就声称工作簿已完成。

| 配方 | 来源与原子动作 | backend | 核验与边界 |
|---|---|---|---|
| 共同底边、等距、组合居中 | Kevin Stratvert《How to Align in PowerPoint》https://www.youtube.com/watch?v=DdHbK2Pl-XI 03:27–06:33 | officecli-native | 参照系先区分页面/选中对象；统一bottom与边到边gap，不等分不同宽度对象中心。真实group保内部布局。3页native通过 |
| 共享网格 | 同片06:33–07:53 | officecli-native | 精确坐标复用边线；未调用GUI智能参考线/吸附/Alt |
| 结论标题+原生图表、负空间、SVG、三分构图 | Leila Gharani《PowerPoint Design Ideas: 5 Tips for Professional Slides》https://www.youtube.com/watch?v=TQiln3CjtvM 00:52–03:53,05:14–08:04 | officecli-native | 5页native；SVG parts保留，两种尺寸清晰；native chart保数值。手工时间线是布局重建，没有调用Designer服务；SVG转原生shape未测 |
| SCQ上下文→A→理由及对应证据 | Analyst Academy《PowerPoint Storytelling: How McKinsey, Bain and BCG create compelling presentations》https://www.youtube.com/watch?v=HnlWiVs9P5o 01:24–02:58,04:50–06:33 | officecli-native | 4页静态演示逻辑/原生可读通过。所有经营数据为虚构教学数据，不能当真实研究。不是每份汇报都强行制造冲突 |
| 圆+三角Union，图片先选后轮廓Intersect | Leila《Merge Shapes in PowerPoint: Custom Shapes and Picture Text》https://www.youtube.com/watch?v=laV-u0KwgVg 02:54–05:19 | hybrid | OfficeCLI基础，COM Union1→Intersect3，真实图片custGeom；标准三角替代源EditPoints微调。原作者照片替换为原创图案 |
| 图片字形Intersect／Subtract／Fragment | 同片05:20–09:07 | hybrid | 图片是PrimaryShape；结果分别1/1/11独立图片几何。保留base文字，结果不再是可编辑词。不是白字盖图或预渲染字形位图 |
| 真正layout图片placeholder | One Skill PPT《PowerPoint Slide Design from Beginner to EXPERT in One Video 🔥100K Special🔥》https://www.youtube.com/watch?v=lxcHLxjkcXQ 12:15–15:07 | hybrid | CLI基础页+COM CustomLayout和placeholderType18，原生画面/结构通过；不把普通图片fill称placeholder。字形placeholder未测 |
| Polaroid框+Pinwheel顺序 | 同片1:08:57–1:18:29 | hybrid | CLI图片shape白边/旋转，COM Pinwheel33、1秒；后3个afterPrevious。首个为自动视频用withPrevious，源首个onClick是公开差别；native7秒六帧逐张旋入通过 |
| 两端点控制柄Bézier | 同片1:33:33–1:40:27 | officecli-native | CLI shape后raw-set原生custGeom，单段cubic；原生导图通过。无四stop渐变线/交叉投影，不是定量Sankey图 |

本增补没有powerpoint-com独立全量生成基础页、也没有external-render的通过配方：保留用户四标签作为路由，不为凑类把native图形转成图像。全套原20页schema通过/issues0，但弧线页25%浅色对比欠佳，已**完全从封装geometry builder排除**，封装版几何7页。

## 参数与防错

- 底边齐：共同bottom=max(y_i+h_i)，再y_i=bottom-h_i。边间距g=(R-L-sum(w_i))/(n-1)。group后移动整体，避免把各子件压到同一个中心。
- 负空间不是必须白底；文本放不下先删减/拆页，不能一味缩字号。三分法是焦点候选；长文本阅读不强制所有段落居中。
- 框架先确认问题与证据，再决定版式。读标题串应能复述主论证，子理由必须支持主结论。
- COM MergeShapes：WithWindow=true打开，先Select目标slide，再Range.MergeShapes；即使有窗，非当前slide也可抛UnexpectedHRESULT。参数从本机PIA核实：Union1/Intersect3/Subtract4/Fragment5。
- 自定义layout应用可能自动实例化placeholder；先查现有Type14/PlaceholderFormat.Type18，重复添加会报maximum placeholders。
- Merge会消费源对象。保留CLI基础文件并保存到新输出，换词回base改，不在转曲后谎称文字可编辑。
- 动态字段成功不代表动画可见。Pinwheel需原生MP4中间帧；静态清晰不能证明动画通过。COM只关闭本次presentation，不Quit共享PowerPoint。

## 可复现命令

所有输出路径由显式参数提供；已存在的PPT/证据目录/验证报告默认拒绝覆盖。

```powershell
python scripts/layout_story_build.py --officecli PATH --outdir NEWDIR --recipe alignment
python scripts/layout_story_build.py --officecli PATH --outdir OTHERDIR --recipe visual
python scripts/layout_story_build.py --officecli PATH --outdir THIRDDIR --recipe story
python scripts/layout_story_geometry.py --officecli PATH --outdir GEOMETRYDIR
powershell -File scripts/layout_story_refine.ps1 -InputDeck GEOMETRYDIR/geometry-base.pptx -OutputDeck NEWFINAL.pptx -PatternImage GEOMETRYDIR/original-pattern.png -EvidenceDirectory NEWEVIDENCE
powershell -File scripts/layout_story_native.ps1 -InputDeck NEWDIR/alignment.pptx -OutputDirectory NEWNATIVE
python scripts/layout_story_verify.py --artifacts NEWDIR --report NEWREPORT.json
```

以上脚本使用Python标准库调用OfficeCLI，不使用python-pptx。

封装后前向验收（2026-11-06）：使用封装`layout_story_build.py --recipe alignment`在全新`validation/layout-story-forward-v1`生成3页；schema通过/issues0，原生3页PNG均已查看，零文字高度溢出，结构独立4检查通过。相同输出目录再次调用以明确错误拒绝，未覆盖原作品。原生输出`validation/layout-story-forward-v1-native`与独立报告`validation/layout-story-forward-v1-check.json`保留。COM已关闭本次presentation。

## 视频自身与组合几何相交

Backend **hybrid**，制作源为仓库`examples/video-intersect`。复用本参考已有可见Normal窗口、当前页选择与Union1/Intersect3；OfficeCLI创建页面、原生视频和poster，COM完成几何与播放绑定，不增加媒体渲染引擎。

1. 先完成自定义轮廓。多个原生部件用Union形成一个真形状；普通group不是相交轮廓。视频应完整覆盖轮廓，并位于轮廓后方。
2. 先选视频，再选轮廓；`Range.MergeShapes(3, video)`明确视频为PrimaryShape。相交消费源对象，基础稿另存，避免丢失可改轮廓。
3. 检查结果仍为Type16/MediaType3，包内有真实videoFile/media关系与`pic/spPr/custGeom`，源视频字节保留。位图poster或矩形视频加孔罩不能代替媒体自身几何。
4. 原生相交可能移除旧自动播放时间轴。本机本例相交后为空；恢复`AnimationSettings.PlaySettings.PlayOnEntry=-1`后，PowerPoint生成一个真正Media Play83。复用该效果、核对Shape.Id并设WithPrevious；不要在已有Play后再盲目AddEffect。已有的其他动画要按实际目标重绑，不由本例推导自动保留。
5. 保存关闭再重开，实际播裸视频对照与裁后视频。检查时间推进、内部内容变化、轮廓外无矩形漏出和结束保持；仅静态导图不能证明媒体在动。

本例三椭圆Union→视频Intersect，原生24秒实播、17个分布帧、第二页187帧边界检查通过。第2页是四文字+一个真媒体对象，无遮罩，嵌入8秒MP4与自产源文件SHA一致。可编辑页面、原生媒体及几何保留；换轮廓/影片后仍须预览，不能把所有编码、GIF、透明度触发或Morph播放并入此范围。

操作参考：[大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)47:55–48:35；本次以Ego Lite复核七张实际操作画面，并用Watch取得该区间英文原始自动字幕，确认视频先选、轮廓后选、Intersect的原生流程。课程的可选渐变描边未纳入本例通过范围。
