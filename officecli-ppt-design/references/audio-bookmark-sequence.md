# 音频书签驱动整组循环

backend: hybrid。OfficeCLI建立原创矢量底稿和自制PCM静音；PowerPoint COM建立原生媒体书签和交互动画。适合多个不同长度的动作按统一节拍重复，并在每轮末尾保留停顿。

来源：[大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)5:03:05–5:05:24。已复读完整对应转写，补看5:00:27、5:04:10、5:04:34、5:05:16四幅实际操作帧；不称连续观看全片。

## 配方

1. 先测整组动作时间。本示范终点2.000秒，静音2.500秒，余0.500秒为停顿。自制WAV没有外来录音。
2. 媒体起点添加CycleStart书签，LoopUntilStopped，自动开始；也可以改成一次点击启动音频。隐藏扬声器，使用嵌入式本地媒体。
3. 创建`TimeLine.InteractiveSequences.Add()`。仅首条用该序列的`AddTriggerEffect(shape,effect,5,audio,'CycleStart',0)`；随后在同一序列`AddEffect(...,WithPrevious)`，用延迟安排动作。本例起点0、0.120、0.200、0.600、1.700秒，整组2秒结束。
4. Exit或Direction会重置Duration，应在这些设置后重新设Duration。自定义位移明确设MotionEffect.Path；只改ToY可能仍保留默认Path。
5. 结构审计必须只有一个直接交互分组，并确认5个效果的延迟、音频长度、2秒终点。再真实放映录屏至少3轮，观察整组重新开始。

不要给每条效果都AddTriggerEffect：实播证实会产生多个顺序分组，每个书签只推进下一分组，不能形成所需整组循环。MainSequence也不支持媒体触发5。原生播放器毫秒回零只证明音频在循环，不能替代动作画面的重复证据。

## 可复现示范

```powershell
python scripts/audio_bookmark_build.py --officecli "E:/ppt skills/tools/officecli.exe" --outdir "E:/work/audio-new/base"
powershell -NoProfile -File scripts/audio_bookmark_polish.ps1 -InputDeck "E:/work/audio-new/base/bookmark-before.pptx" -OutputDeck "E:/work/audio-new/bookmark.pptx" -Evidence "E:/work/audio-new/native"
python scripts/audio_bookmark_verify.py "E:/work/audio-new/bookmark.pptx" --report "E:/work/audio-new/audit.json"
powershell -NoProfile -File scripts/audio_bookmark_record.ps1 -Deck "E:/work/audio-new/bookmark.pptx" -Evidence "E:/work/audio-new/play" -Ffmpeg "C:/path/to/ffmpeg.exe"
```

全部路径显式指定且拒绝覆盖；不退出共享PowerPoint应用。原生窗口需串行操作。录屏先关闭演讲者视图，用EnumWindows定位screenClass并裁切该窗口所在桌面区域；硬件加速窗口不能只依赖PrintWindow静态抓取。

## 已验与边界

研究v6：OfficeCLI校验零错、原生静态PNG、真实10秒桌面录屏。40帧对照可见多轮闪光→照片进入/移动/退出→停顿；像素计时闪光1.467/4.067/6.567/9.067秒，稳定间隔2.500秒，首次间隔多0.100秒为此次采集调度抖动。5效果一分组、2秒终点、2.5秒嵌入PCM均另行审计。

只通过本原创矢量示范的多效果序列循环。Ink Replay五秒绘制+两秒停顿有[独立配方](ink-replay-pause.md)，不能由这里代证；在线音频、跨页音乐、Mac或网页版仍未覆盖。改变音频时长后必须保证动作终点早于下一次书签，并重放验证；不要把本示例的节拍机械套进更长动作。

证据：`distill/audio-bookmark/source-notes.md`、`after-v6.pptx`、`native-v6/`、`play-v6/native.mp4`及`analysis.json`。原始作者素材不随技能打包。

封装后前向：`verification/audio-bookmark-forward/`。直接运行包内四脚本生成新底稿、新成品和新10秒桌面影片；CLI validate零错、单交互分组/5效果审计通过、40帧接触表已目视。检测四次闪光1.467/4.067/6.600/9.100秒，后两间隔2.533/2.500秒；照片四次进入后退出，符合统一时钟，帧采样和录制存在约一至三帧抖动，不宣称无误差帧同步。

录屏起点说明：脚本先Run放映（音频自动开始），等待1000毫秒，再定位窗口/启动ffmpeg；没有重置音频。因此影片0秒不等于放映0秒，0–0.5秒已有照片是被截断的首轮尾段。验收使用后续完整周期，不把影片首帧当放映初始状态；此证据未单独验证放映启动瞬间的首帧隐藏。

官方接口：[AddTriggerEffect](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.addtriggereffect)、[MediaBookmarks.Add](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.mediabookmarks.add)。
