# 原生透明 GIF 动态纹理

[纹理演示PPTX](gif-texture.pptx) · [可编辑动画源PPTX](animation-source.pptx) · [原生播放](playback.mp4)

原创TRACE双页小样：点击首页按钮进入同页对照，左侧为PowerPoint原生导出的透明GIF，右侧为同一GIF与原生花瓶形状真正Intersect后的活动纹理。没有视频或外部编码GIF替代，轮廓与底色可编辑；逐帧纹理需要从另附的原生动画源重新导出。

动画源由五个原生方框与青色相位点组成；整组Spin90°和Grow130%，各2秒、自动翻转、平滑开始/结束各1秒。通过PowerPoint真实“导出→创建动态GIF”导出特大1080p、透明背景、0秒自动时长、仅源页。最终GIF为1920×1080、99帧、98个不同画面、编码周期3.96秒，四角逐帧透明。导出后原来的质量、透明与范围设置均已恢复。

本机直接相交会使裁后GIF静止。修复保留原生Intersect生成的custGeom，把srcRect等价改成负向stretch/fillRect；不改GIF字节、对象位置或裁形路径，也不增加外部遮罩。去掉srcRect却不补偿填充会挤压纹理，因此不能当成正确修复。

发布源全新制作后，最终PPTX经PowerPoint再次保存、关闭、重开，原生18秒实播并复看19个分布帧。225个裁后采样帧中，轮廓外4像素抗锯齿容差以外无白纹理；与裸GIF对应区域在1像素容差下的最低匹配约99.99%，画面比例保持。最后一段仍明显运动，覆盖三个以上完整编码周期。原GIF字节与原生相交几何前后不变，OfficeCLI验证通过。

在仓库根目录运行，使用新目录及新输出：

```powershell
python officecli-ppt-design/scripts/native_gif_build.py --officecli OFFICECLI --stage source --out NEW/source-base.pptx
powershell -File officecli-ppt-design/scripts/native_gif_finish.ps1 -Stage source -InputDeck NEW/source-base.pptx -OutputDeck NEW/animation-source.pptx -Native NEW/source-native
powershell -File officecli-ppt-design/scripts/native_gif_export.ps1 -Deck NEW/animation-source.pptx -Gif NEW/native-texture.gif -Evidence NEW/export-native
python officecli-ppt-design/scripts/native_gif_build.py --officecli OFFICECLI --stage texture --gif NEW/native-texture.gif --out NEW/texture-base.pptx
powershell -File officecli-ppt-design/scripts/native_gif_finish.ps1 -Stage intersect -InputDeck NEW/texture-base.pptx -OutputDeck NEW/intersect.pptx -Native NEW/intersect-native
python officecli-ppt-design/scripts/native_gif_repair.py --input NEW/intersect.pptx --out NEW/repaired.pptx
powershell -File officecli-ppt-design/scripts/native_gif_finish.ps1 -Stage save -InputDeck NEW/repaired.pptx -OutputDeck NEW/gif-texture.pptx -Native NEW/final-native
```

需要Windows桌面PowerPoint、Windows PowerShell、Python与OfficeCLI；字体为Noto Serif SC/Noto Sans SC。自动导出脚本按本机中文界面操作，其他语言可在相同原生面板手动导出。质量下拉的ValuePattern只读回，实际选择使用可见选项，脚本会恢复原设置并检查自己的保存对话框已关闭。

修复限本例正值、非空srcRect与无既有fillRect的图片；不要覆盖别的复杂填充。新的轮廓、GIF或软件环境仍需实际看几轮；本例未测试GIF指定对象触发、Wipe遮罩、任意滤镜或实时流体。
