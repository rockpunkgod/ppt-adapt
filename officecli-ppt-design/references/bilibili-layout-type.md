# 中文文字、字体与多图版式

本模块来自五个达到播放量门槛的B站短片。以完整可用字幕和操作帧还原方法；样例更换了文案、风景和本机字体，属于机制复现。静态版式和第7页顺序入场有本机PowerPoint验证，不将源片所有装饰动画计为通过。

## 先按信息选择版式

- 长段正文保持稳定字形和字重；标题/数据再以粗细、大小和主题色建立主次。中文、拉丁字母的同一pt字号不代表相同视觉高度，需要并排看实际字面。
- 纯文字先统一轴线，再拉开段间距离，再强调结论。不要靠空格凑对齐，也不要把小号长段文字居中排成参差形状。重复纹理属于装饰，主信息仍保持清楚。
- 中文短标题可用闭合修饰连成视觉块，也可拆成少量文本对象做轻微高低错落，并在大字空隙穿插小字。长正文使用清晰字形、分组和留白，不照搬英文占位文字的密度。
- 多图先明确主题区域与图片数量，再选瀑布流、轻微倾斜、平铺蒙版、原生形状窗口、环绕主题或字形图片填充。不能用压扁图像来凑尺寸。

## 原子配方与backend

|配方|来源与时段|backend|实作与验收|
|---|---|---|---|
|字重/字形/字面高度核查|BV1tz4y1C7PJ 00:49–02:51|officecli-native|第1页本机微软雅黑/Arial示例；native检查断句、主次、字号实际高度|
|轴线、留白、主次与主题色|BV19V411U7hs 02:15–07:02|officecli-native|第2页；坐标、段间距和字号分别控制，不靠空格|
|文字重复纹理与真字形描边|同片07:03–09:19|hybrid|CLI放置24个原生文字对象；COM Font.Fill.Visible=0与Font.Line控制字形线条；第3页|
|中文标题闭合成块|BV1TohbzdEAL 01:04–01:39|officecli-native|第4页框与文字分离；矩形只是一种闭合形式，不是所有页面必须加框|
|短标题节奏与小字穿插|同片01:47–02:13|officecli-native|第5页短标题两对象错落；保留正常字形比例，不将长正文随机错位|
|中文视觉密度重组|同片02:17–03:27|officecli-native|第6页按语义拆四块，统一轴线与字号层级；示例文案原创|
|整图+1行5列表格+交替透明度|BV1T54y1c7EU 00:21–01:59|hybrid|第7页是真table，格2/4的透明度55%，其余45%，即OOXML alpha45%/55%；保留内竖线|
|序号浮入、整段文字飞入、逐列推进|同片01:06–01:44|hybrid|序号一个对象；标题与正文同一文本对象。COM把数字EffectType修为原生Float=30，10个Effect顺序推进；看中间帧确认组内不穿插|
|瀑布流与轻微倾斜|BV1SY411J732 00:07–00:18|officecli-native|第8/9页固定列宽、间隙和裁切；只通过静态布局，不包含源片瀑布流运动|
|图片平铺+暗色蒙版|同片00:18–00:23|officecli-native|第10页平铺图片上覆原生暗色形状；不是背景模糊|
|矩形/六边形/三角形真窗口|同片00:23–00:33|officecli-native|第11页prstGeom+image fill+srcRect；画布外用背景色遮挡不等同真实裁切|
|图片环绕主题|同片00:33–00:37|officecli-native|第12页主题区独立，图片不压住主要结论|
|字形内部填图片|同片00:37–00:41|hybrid|第13页每字母独立原生文字，COM Font.Fill.UserPicture；不是文字框背景图。原生UserPicture使用stretch；若主体比例重要，应先按字面窗口准备合适素材并复测|

原子COM操作（文字描边、单元格透明度、EffectType修正、字形Fill.UserPicture）本身属于 `powerpoint-com`；组合到CLI基础页后是 `hybrid`。本批没有需要 `external-render` 的已测配方。

## 直接运行

```text
python scripts/bili_layout_build.py --help
python scripts/bili_layout_build.py --officecli officecli --out ./new/layout-base.pptx --landscape /path/wide.jpg --portrait /path/portrait.jpg
powershell -NoProfile -File scripts/bili_layout_polish.ps1 -Deck ./new/layout-base.pptx -Output ./new/layout-final.pptx -Landscape /path/wide.jpg -Portrait /path/portrait.jpg -RenderDirectory ./new/native
python scripts/bili_layout_verify.py ./new/layout-final.pptx --report ./new/structure.json
```

输出PPT、精修PPT、报告和渲染目录都应选新路径。脚本拒绝覆盖。输入为调用方有权使用的JPEG；不捆绑视频截图、作者模板、字体文件或验证照片。基础PPT是13页可编辑对象，图片作为独立填充素材，不是整页截图。

独立检查验证24个原生字形描边、原生1×5表格与alpha、10个正确目标/类型动画、4个字形图片填充和30个shape图像的裁切比例。它不单独证明动效；仍需按总技能门槛看原生PNG和影片第7页入场前中后。

## 已从原生作品修正的问题

1. 多行框过矮会截字；初版结构检查发现后增高。字体卡片出现句号单独换行，改为有语义的短行，不用缩小正文掩盖。
2. 第7页把标题/正文分成两份Fly，即使同拍也可能在中间态互穿。改为同一文本对象，COM仅改变首段28pt/加粗，正文20pt；移动中保持相对位置。
3. 本机OfficeCLI 1.0.160的`effect=float`回读为float，但实际presetId=12。PowerPoint本机PIA定义12=Peek、30=Float，不能根据命令名称判定效果；精修脚本按对象名把EffectType设为30并重设0.4秒/顺序触发。Fly仍为2。
4. `Transparency=.45`表示45%透明，OOXML alpha是55000。不要把透明度和不透明度直接照抄。深灰遮罩是本作对不同照片的适配；换图后仍需查正文对比。

## 来源和范围

播放快照为2026-11-06本地00:35–00:36：旁门左道PPT字体片181,427；SeeOneQF纯文字片179,035；乔乔乔ppt表格图文片29,878；利兄日志PPT多图片156,180；旁门左道PPT中文片277,329。上述均≥10,000。

字体片04:00前后的六组屏幕示例已核对：思源黑体CN Bold/Gotham Bold、思源黑体CN Regular/Bahnschrift、汉仪旗黑-75W/Soho Gothic、方正风雅宋/Didot、思源宋体CN Semi/Georgia Pro、微软雅黑/Arial。这是源画面记录，不是这些字体一律可商用或必须安装的清单；本样例用本机已有字体验证方法。

BV1UH4y127CS（1,650）和BV12GcxzPEnb（5,573）按用户门槛跳过，不从标题补写层级/密集时间轴能力。两条长课分P由独立笔记处理，不并入这五短片的通过范围。

来源：[字体搭配](https://www.bilibili.com/video/BV1tz4y1C7PJ/)、[纯文字排版](https://www.bilibili.com/video/BV19V411U7hs/)、[表格图文](https://www.bilibili.com/video/BV1T54y1c7EU/)、[多图](https://www.bilibili.com/video/BV1SY411J732/)、[中文排版](https://www.bilibili.com/video/BV1TohbzdEAL/)。
