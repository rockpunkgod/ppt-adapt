# 一个真实文本框，按段落、字符或单词出现

backend：`hybrid`。OfficeCLI制作保留原文字的页面，PowerPoint COM建立原生文本范围动画。不是把文本拆成许多独立形状。

## 来源与方法

[BV1VTDGYsEmd P21](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=21)2049–2175秒：同一框里的三个自然段逐段出现，回车形成段落。已核2049/2070/2090/2110/2140/2175秒整帧和相应OCR。[One Skill大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E&t=20067s)5:34:27–5:35:24：Swivel，标题by letter、正文by word，单元时长0.2秒、字间延迟10%。Ego Lite新增20080秒完整菜单帧确认By letter/10%，20105秒确认By word选择，20108秒可见正文局部分组旋转；不是只依据自动字幕猜预设。

当前原创样稿将三种机制放在三页，所有目标文字仍各为一个可编辑文本框。段落页使用三个自然段/三次点击；字符页为英文标题、单词页为英文句子，各一次点击。保留源Swivel19/200ms/10%参数；未复刻源背景/字体，也未声称源标题WithPrevious→正文AfterPrevious整套时序已通过。

## 原生实现

- 段落：`MainSequence.AddEffect(shape, Fade10, level1, OnPageClick1)`生成三条真实paragraph范围；明确将每条Timing.TriggerType设1、Duration设0.5，保存后范围为0/1/2且各是clickEffect。默认AddEffect曾把后两条设WithPrevious；不能只检查有三条效果就宣称三次点击。
- 字符/单词：先AddEffect(Swivel19)，再`ConvertToTextUnitEffect(effect,1或2)`，**使用返回的新Effect**，设Duration0.2。存为原生`p:iterate type=lt/wd`、`tmPct=10000`。不拆字形，不手工堆多个淡入副本。
- 普通页面点击须可推进此样稿的时间轴。本机对照中AdvanceOnClick=false时物理点击后GetClickIndex仍0；相同精修仅改为true后按预期推进。不要复制Pin/Kiosk门控设置来阻断普通文本构建。定时翻页关闭；效果播完再点可能进入下一页，属于普通放映行为。

## 可复现入口

基础构建/结构检查只需Python标准库和同包`build_components.py`。精修与播放需Windows PowerPoint；真实录屏需FFmpeg。像素分析另用已有Pillow/NumPy，未用于创建PPT。各输出新路径、拒覆盖。

```powershell
python scripts/text_build_build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/text_build_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Evidence NEW/native
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/text_build_record.ps1 -Deck NEW/final.pptx -Directory NEW/live -Ffmpeg FFMPEG
PY_PIXELS scripts/text_build_measure.py --directory NEW/live --ffmpeg FFMPEG --ffprobe FFPROBE --report NEW/measure.json
python scripts/text_build_verify.py --base NEW/base.pptx --deck NEW/final.pptx --measure NEW/measure.json --report NEW/checks.json
```

COM精修只接受该三页全新基础稿、目标名text-build-target且无原动画，不负责替换用户现有整页时间轴。录屏以COM GotoSlide选择测试页，随后五次真实鼠标点击推进效果，逐次核实际页与ClickIndex；不把COM换页当文本点击通过。采用自己可见放映窗口、前台/命中/实际指针坐标守卫，DWM物理区域记录；仅关自有文稿，不退出共享PowerPoint。

## 前向结果与量测修正

`validation/forward-text-build/`：包内全新CLI三页→COM→三张原生PNG→14秒真实放映→34项结构/状态检查→OfficeCLI validate全部通过。实际502个独立VFR帧（请求60fps不等于60个独立帧/秒），采用FFprobe真实PTS和passthrough解码，无补帧。

四张段落状态分别空/一段/两段/三段，原始文本完整保留；字符页12个片刻呈前字符已出现而末字符未出现，首末检测相隔0.300秒；单词页3个片刻呈首词已出现而末词未出现，相隔约0.083秒，随后全部稳定。后两者为像素检测时差，不伪称精确的单元延迟测量；200ms与10%另由保存结构核对。

早期像素分析曾把COM边界扩展后的上一行尾部算进下一段，也把H264色度抽样/缩放边缘损耗误判为字符未全到位。修正为从完整终态图中独立分离三条文字行，并只量测终态色块向内侵蚀1像素后的实色核心；保留原阈值，使用相同前向脚本重测。旧失败报告保留，未放宽阈值来凑通过。

范围：当前英文字符/单词、三个自然段、Arial与本机原生Swivel/Fade；未测中文分词、复杂连字、双向文字、任意字体替换、任意段落层级或触发组。文字、字号、布局、颜色或节奏改变后，重建并重新原生播放/校准像素区域。分析器是固定样稿的验证工具，不是通用OCR或任意文本跟踪器。
