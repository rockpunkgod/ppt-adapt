# 母版与布局的独立动画时间轴

backend：`officecli-native`。OfficeCLI制作全部形状、页面、本地Fade与真实master/layout timing；COM只读计数、导图、原生放映和录屏，不进行authoring。不是把继承形状复制到普通页。

来源：[完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E&t=3838s)1:03:58–1:06:32，布局专属直线路径/AutoReverse/Until End；[27P P22](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=22)2643–2802s，普通页不可见的母版时间轴、2秒光晕和9秒背景移动重复到页末。全文可用转写/OCR与2727s真实until-end对话框已核。未复制狐狸照片、电路图或作者PPT。

设计规则：需要全模板共有的运动放master，需要分支专属的运动放对应CustomLayout，逐页讲述的点击动画留在slide。三个部件各有自己的timing和目标ID；不要把普通页动画窗格为空误判为无动画。动画继承与静态Logo继承是两种能力。

受测原创：A/A/B三页，共同master橙点水平往返，layout A青方块向下往返、B紫方块向上往返；前向1秒+AutoReverse1秒=2秒周期，repeat到页末。每页只有三个local形状与一个0.4秒onClick Fade蓝块。B的空Title/Subtitle布局槽不是动画对象，不产生普通页副本。

## 包内复现

所有路径使用绝对路径，输出目录必须不存在。builder/verifier只用Python标准库；像素量测需Pillow/NumPy、FFmpeg/FFprobe；原生验证需要本机PowerPoint和独占UI。不会Quit共享PowerPoint，不访问ChartData。

```powershell
python scripts/master_animation_build.py --officecli OFFICECLI --out NEW
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/master_animation_probe.ps1 -Deck NEW/master-loop.pptx -Evidence NEW/native
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/master_animation_record.ps1 -Deck NEW/master-loop.pptx -Evidence NEW/live -Ffmpeg FFMPEG
PY_PIXELS scripts/master_animation_measure.py --live NEW/live --ffmpeg FFMPEG --ffprobe FFPROBE
python scripts/master_animation_verify.py NEW
```

CLI高层animation当前拒绝master/layout路径，且高层Add的repeat/autoReverse曾被忽略。脚本先用OfficeCLI生成自己的合法carrier timing，然后用OfficeCLI raw-set写入`/slideMaster[1]`、`/slideLayout[1/2]`；这两个raw部件名大小写与高层路径不同。**所有spid引用包括bldP都必须映射到目标部件真实shape ID**。effect cTn显式写`repeatCount=indefinite`、`autoRev=1`，不能把tmRoot的dur=indefinite误当具体效果在循环。

## 验收边界

原生Master/两CustomLayout/三Slide的MainSequence分别仅一效果，普通页只有local动画；父部件的对象不出现在slide shape tree。UntilEnd在本机COM读回RepeatCount=-2147483648，要结合XML/真实播放，不拿负数自行推算次数。

18秒真实录屏：首A1驻留>7.2秒，覆盖三个2秒周期且继续，物理点击本地Fade不打断继承运动；后两次物理空白click自然到A2/B3，两个继承时间轴重新起相且A/B互不显示。末段**明确调用GotoSlide(1,ResetSlide=true)**，仅证明显式重置，不能外推默认后退/返回行为。

量测使用FFprobe实际PTS和passthrough帧，不补帧；标题区分真实页段，用已知纯背景校准投影窗口黑边后的画布。三角波拟合按已知运动方向消除正负幅度/半周期同解。固定标准：周期2±0.05秒、幅度误差<3%、位置RMSE<6px、重进页起相误差<0.2秒；相位初始值不假设和录屏启动同刻。local Fade隐藏/显现、A/B互斥也须检查，不能仅看结构。

不外推全部模板/占位符动画、任意长后台运行、跨版本PPT、默认返回状态、母版交互触发；图案/路径/时长改动后重录并重新计算。首次研究的raw路径错误、漏bldP、Add忽略repeat负例都已保留；分析器修正画布黑边和符号歧义后仍按同一录屏/同一阈值重算，不冒称首次量测通过。

最终包内fresh前向：`validation/forward-master-animation/README.md`，真正全新carrier/三页PPT、原生3PNG、18秒真实桌面/72张分布帧目视与478实际PTS帧量测全部通过。四页段周期2.000/2.005/2.000/2.000秒，RMSE≤3.70px，重入起相误差约25/0/10ms；A/B互斥、本地Fade和三轮后继续/自然跨页/显式Reset均过。原稿SHA、拒覆盖、schema通过，不复用研究成品。
