# 同一数值驱动半环图表、真实支点指针与滚动数字

backend：`officecli-native`。OfficeCLI创建真实doughnut、原生形状/组/文本、Spin/路径及原生timing；COM只读原生导图、TimeLine读回和实际播放验证。图表使用numLit，不宣称PowerPoint“编辑数据”工作簿可用；不调用ChartData、Series.Values/Formula或Shape.Chart API。

来源：[完整大师课9Ewb0Rd3P3E](https://www.youtube.com/watch?v=9Ewb0Rd3P3E&t=21692s)6:01:32–6:27:26。已读全段可用转写，并看6:06:42数据窗口、6:17:06下半弧/左指针、6:22:18数字路径与最终100%分布帧。6:22:18还是Down方向的中间态，6:22:27才修Up；不混拼源中50/75/100为同一终值。

## 核心映射

每页从一个整数v生成：图表两项`[v,200-v]`，角度`360*v/200 = 1.8*v`，十位`v//10`、个位`v%10`及各自滚动距离。默认原创25/60，对应45°/108°。两页各从00开始，滚字是装饰过渡，**不是每帧实时计量值**。

1. 真doughnut，hole80%，第二数据点noFill。`colorrule=100:26BFAF:none`在本两项1..99输入中生成第一点青色、第二点真正a:noFill。不是把隐藏扇形涂白。无legend/title，manual plot填满正方形以确保旋转中心。
2. 最终首切片270°显示上弧；初态90°配Spin180°。下半annulus罩遮住转入前的弧，内孔保留给数字；遮罩径向扩大0.2%并上移0.02cm消除抗锯齿漏线。
3. 针、hub与直径14cm透明envelope组成真group，group中心等于chart中心；针从左水平起，原生Spin到1.8*v。不能靠画一条终态假针代替实际绕共同支点转动。
4. 两个真实单文本框，各行固定45.35pt；十位0..9，个位0..9+0。十位路径2秒到目标；个位0.25秒完整滚动×4，1秒延迟后以1秒路径到目标。上下窄罩隐藏其余行。60的个位0快转后保持0，十位仍在主时序完成。
5. chart/needle/tens主Duration2s；needle accel0.375/decel0.5、chart/tens decel0.5；最后ones Delay1s/Duration1s。首个chart onClick，其余withPrevious，同一次真实点击启动。

## 五脚本入口

所有路径使用绝对路径，输出新目录、拒覆盖。builder只用Python标准库；检查需本机Pillow/NumPy与Windows Arial Bold字体文件；原生录屏需FFmpeg/FFprobe及Windows PowerPoint。不能并发操作共享PowerPoint UI，不Quit，仅关闭自己打开的稿件。

```powershell
python scripts/gauge_sync_build.py --officecli OFFICECLI --out NEW --values 25,60
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/gauge_sync_probe.ps1 -Directory NEW -Evidence NEW/native
PY_PIXELS scripts/gauge_sync_pixels.py --directory NEW --native NEW/native
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/gauge_sync_record.ps1 -Deck NEW/gauge.pptx -Evidence NEW/live -Ffmpeg FFMPEG
PY_PIXELS scripts/gauge_sync_measure.py --directory NEW --ffmpeg FFMPEG --ffprobe FFPROBE
```

builder产出动画gauge.pptx与独立静态终态final-reference.pptx；后者仅为比对参考，不能代替播放验收。原生script先Run并启动编码器，在00静态保持后才物理点击，首段不因录屏启动而丢动画。第二次空白物理click自然进第二页，再真实click触发。

## 已验证的必要修正

- Motion效果outer cTn的fill=hold还不够：**tens和ones收尾的child cBhvr/cTn同样必须hold**，快转ones用remove。原始v3曾瞬间到25/60后回00；真实录屏与独立字形检查拒绝该版本。只截正确的一帧不能验收保持。
- CLI给末ones的Delay1s曾写到祖先wrapper，而COM Timing读0；不能仅凭getter断言运行错误。当前脚本把同一延迟放effect cTn、祖先清0，原生读1s；v1层级差异原样保留。
- 原生Timing2s/Delay1s是参数证据；实屏时序另验触发前00、1.5s尚非终值、2.1s已终值且持续。不会把较粗屏幕边界写成精确测得2.000s。分段匀加速假设拟合曾误推1.84s，已明确弃用，原片与失败报告保留。

结构检查读真实chart cache、noFill、group支点、Spin角度、文本列、路径与native5效果；像素独立扫弧、测针、用Arial Bold十字形词典识别真实数字，最终每位glyph IoU>0.85、角误差<0.6°。视频按FFprobe实际PTS/passthrough解码，画布先排除投影窗口黑边。快速滚动另检查窗口像素变化；终态后多帧不得回00。

已测默认25/60、Arial Bold40pt和45.35pt行距；输入接受两个1..99整数，其余值/字体/字号/画布需重新原生验收。100需三列，未在本配方验证。省略源白刻线、按钮、comment bubble，不外推完整视觉模板或每中间帧三个部件都显示同一个实时值。

包内exact前向：`validation/forward-gauge-sync/README.md`。从零生成、4原生PNG、11秒真实三click/66帧分布目视、308实际PTS帧、结构/弧角/真支点/字形/保持与两稿SHA/拒覆盖/schema通过。原生Timing2s/Delay1s独立读回；实屏仅称1.5s尚非终值、2.1s已完成并保持，不能称精确实测2.000s。晚期5/47帧继续验证25/60，原先回00反例被拒。

即时advance日志只作物理输入时钟，可仍回报离开页；到达必须另验next=2与真实画面切段。首次分析器把即时快照当已经到页2而早拒，修正语义后同片重算，未改导航/未重拍；此分析器假设修正与PPT本身故障要分开记录。
