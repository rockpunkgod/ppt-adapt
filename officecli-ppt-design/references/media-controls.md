# 原生媒体裁剪、音频淡化与书签定位

| 能力 | backend | 运行方式 |
|---|---|---|
|保留原始媒体、裁掉前后部分|officecli-native|OfficeCLI媒体trimStart/trimEnd；不预剪源文件。|
|声音渐强、渐弱|hybrid|OfficeCLI嵌入原始音频；COM设MediaFormat.FadeInDuration/FadeOutDuration。|
|定位媒体关键时间|hybrid|COM创建MediaBookmarks，原生Player.GoToNextBookmark/GoToPreviousBookmark。|

来源：[27P课程P19／片内第18课](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=19)，738–828秒裁剪/书签，1434–1476秒音频淡化。精读每3秒OCR并核看747、768、795、1443、1467秒操作帧，不称连续逐秒观看。1443秒画面只确认淡入5.00、淡出0.00，未把口述意图当成两个已设5秒。

## 参数与操作

**trimEnd是尾部裁掉时长，不是保留结束点。** 原媒体12秒保留2–10秒，应传trimStart2s、trimEnd2s。公式为`尾裁长度=原始长度−保留终点`；原生COM对应StartPoint2000、EndPoint10000，XML则是st2000/end2000。把trimEnd设成10s会连同前裁2s移除整段，本机原生曾因此重置成0–12秒；不要把这个错误参数例子称作CLI不支持。

淡化用原生MediaFormat的两个毫秒属性，本例各2000。源WAV保持恒幅，没有在FFmpeg预做渐强渐弱。书签在未裁剪视频6000/9000毫秒处；通过Player原生Next、Next、Previous观察6→9→6秒。没有用CurrentPosition赋值或翻页冒充书签定位，也不是此前的书签触发形状循环。

`media_controls_build.py`从头生成带源秒号的12秒视频、逐秒音阶及三页OfficeCLI底稿；自制媒体只用来测量。`media_controls_polish.ps1`必须先回读确认CLI裁剪正确，随后只设fade和书签，不能悄悄用COM修trim。

## 原生音画验证

运行环境为Windows x64、已安装本机PowerPoint与PIA、PowerShell7；需要已有FFmpeg。录音器由PowerShell Add-Type编译包内C#，调用Windows官方应用进程loopback接口；不安装组件、不打开麦克风、不切默认设备、不改全局音量。该接口需要Windows build20348或以上，本机26300已测。

**录音按PowerPoint进程树隔离，不按单个文稿隔离。** 原生验证时只能播放本次自有媒体，不能一边录一边让同一PowerPoint的其他文稿播放声音。脚本从自己的PowerPoint窗口解析PID，拒绝已有其他放映，只关闭自己的文稿/放映，保留其他窗口。

```powershell
python scripts/media_controls_build.py --officecli "C:/path/officecli.exe" --ffmpeg "C:/path/ffmpeg.exe" --outdir "E:/work/media-new/fresh"
pwsh -NoProfile -File scripts/media_controls_polish.ps1 -InputDeck "E:/work/media-new/fresh/base/media-before.pptx" -OutputDeck "E:/work/media-new/media.pptx" -Evidence "E:/work/media-new/native"
pwsh -NoProfile -File scripts/media_controls_record.ps1 -Deck "E:/work/media-new/media.pptx" -Slide 1 -Evidence "E:/work/media-new/captures/trim" -Ffmpeg "C:/path/ffmpeg.exe" -Seconds 9
pwsh -NoProfile -File scripts/media_controls_record.ps1 -Deck "E:/work/media-new/media.pptx" -Slide 2 -Evidence "E:/work/media-new/captures/fade" -Ffmpeg "C:/path/ffmpeg.exe" -Seconds 9
pwsh -NoProfile -File scripts/media_controls_record.ps1 -Deck "E:/work/media-new/media.pptx" -Slide 3 -Evidence "E:/work/media-new/captures/bookmarks" -Ffmpeg "C:/path/ffmpeg.exe" -Seconds 8
pwsh -NoProfile -File scripts/media_controls_control.ps1 -InputDeck "E:/work/media-new/media.pptx" -OutputDeck "E:/work/media-new/control.pptx" -Evidence "E:/work/media-new/control-native"
pwsh -NoProfile -File scripts/media_controls_record.ps1 -Deck "E:/work/media-new/control.pptx" -Slide 2 -Evidence "E:/work/media-new/captures/control" -Ffmpeg "C:/path/ffmpeg.exe" -Seconds 9
python scripts/media_controls_verify.py "E:/work/media-new/media.pptx" --assets "E:/work/media-new/fresh/assets" --native "E:/work/media-new/native" --captures "E:/work/media-new/captures" --report "E:/work/media-new/structure.json"
python scripts/media_controls_audio_bounds.py "E:/work/media-new/captures/trim/powerpoint-audio.wav" --report "E:/work/media-new/trim-audio.json"
python scripts/media_controls_fade_analyze.py --fade-deck "E:/work/media-new/media.pptx" --control-deck "E:/work/media-new/control.pptx" --fade-wave "E:/work/media-new/captures/fade/powerpoint-audio.wav" --control-wave "E:/work/media-new/captures/control/powerpoint-audio.wav" --report "E:/work/media-new/fade-audio.json"
```

全部输出显式指定、拒绝覆盖。录音先于Run，桌面影片晚于Run及编码器启动；独立记录时间，不宣称帧精确同步。应看真实影片源秒号与书签跳变，并检查WAV，而非只看参数回读或CreateVideo导出音轨。

## 测量边界与证据

变化音高的实际输出RMS可能带有共同的播放链频率响应，不能假定它在每秒都同强度。早期未经校准的报告false已保留；严格对照须使用**同一嵌入WAV、相同codec/trim/80%音量/时序/Player/采音链，仅fade=0**。先按音高转折对齐源时间，再比幅度；不通过放宽原阈值制造通过。MP4 AAC不能直接充当WAV的严格同媒体控制。

研究原型已验证：保留段有效音约8.10秒，首D4、末A4；书签原生位置6/9/6秒，暂停显示帧约提前67毫秒，不称帧准确。严格同WAV对照后中段归一增益约.992–1.003，前段.091→.847、末段.901→.203，原淡化判据通过。三段和控制均无采音丢包/时间戳错误。媒体嵌入字节与原始自制素材完全一致，未破坏性剪切或烘焙fade。

书签使用未裁剪视频单独测试；不由此推断裁剪后书签偏移规则。片段Player结束位置约8.02秒，而MediaFormat保留终点是原源10秒，勿混用坐标。换素材、时长或目标环境应重新原生验证。

**包内fresh前向已通过**：`verification/media-controls-forward/`由包内脚本重新生成素材/三页/成品/控制稿；schema零错、15项独立审计全真、全部媒体字节等于新生成原始素材。四段实际录音绑定同PID与相同recorder/C#源哈希，无丢包或时间戳错。保留段实测8.05秒、首D4末A4；书签事件6/9/6与GDI源秒号已目视。严格同WAV控制后中段增益.993–1.011，前段.087→.841，末段1.011→.205，原淡化判据通过。

本配方通过的是trim、音频fade和未裁剪视频的bookmark定位。**跨页播放与停止生命周期尚未验证，不能将整个媒体族标完成。**

官方：[MediaFormat.StartPoint](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.mediaformat.startpoint)、[EndPoint](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.mediaformat.endpoint)、[FadeInDuration](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.mediaformat.fadeinduration)、[FadeOutDuration](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.mediaformat.fadeoutduration)、[GoToNextBookmark](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.player.gotonextbookmark)、[GoToPreviousBookmark](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.player.gotopreviousbookmark)、[微软ApplicationLoopback样例](https://github.com/microsoft/Windows-classic-samples/tree/main/Samples/ApplicationLoopback)。
