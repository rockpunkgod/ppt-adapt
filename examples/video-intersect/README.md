# 原生视频裁形

[可编辑PPTX](video-intersect.pptx) · [原生播放](playback.mp4)

两页TRACE视觉技术小样：第1页为同一视频的裸矩形对照，第2页为三个原生椭圆Union后与视频Intersect得到的自定义轮廓。8秒动态色标为本例自产诊断素材，方便识别画面变化、时钟、真实边界和终帧保持。

视频本身保留原生媒体对象（Type16、MediaType3）和custGeom；第2页仅有四个文字对象与一个视频，没有叠加遮罩。裁前裁后嵌入MP4字节相同。PowerPoint相交会移除旧自动播放时间轴；本例恢复PlayOnEntry并复用它生成的唯一Media Play效果，设置WithPrevious。不要再追加第二个Play。

发布源全新生成后保存、关闭、重开并在本机PowerPoint播放24秒；11秒处物理Right进入第2页。复看17个分布帧：裸片在动、裁后时钟0.333→7.900秒且画面在动、轮廓稳定、最后保持。以15fps检查第二页187帧，轮廓外4像素抗锯齿容差以外未检出视频画面；包内媒体哈希、custGeom、实际媒体目标与全部动画绑定一致，OfficeCLI验证通过。

在仓库根目录运行，使用新目录：

```powershell
python examples/video-intersect/build.py --officecli OFFICECLI --ffmpeg FFMPEG --out NEW/base.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File examples/video-intersect/finish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Native NEW/native
```

需要Windows桌面PowerPoint、Windows PowerShell、FFmpeg，标题/正文使用Noto Serif SC/Noto Sans SC。builder在输出目录的assets子目录生成自有测试媒体；最终PPTX已嵌入媒体，可独立打开。

本例只验证三椭圆联合轮廓、H.264 MP4、自动播放及结束保持。新的轮廓、编码、影片替换或复杂效果需重新预览；不将此结果扩大为GIF相交、透明视频、指定对象触发组合或Morph中的实时播放。
