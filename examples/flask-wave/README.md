# 瓶形液面与双层波浪

[可编辑PPTX](flask-wave.pptx) · [原生播放](playback.mp4)

两页原创TRACE视觉小样，使用同一瓶形展示两种动效。第1页单击上升、末段回弹；再次单击进入第2页，再单击让两层波浪沿不同斜线升起。文字、瓶形、液体与遮罩均为原生可编辑对象，不使用预渲染视频代替效果。

复用OfficeCLI自定义曲线、原生Subtract、幻灯片背景填充、Fly/Bounce及COM路径动画。全画布矩形减去瓶形产生真孔；上层另放无填充白边瓶形，背景填充遮罩挡住瓶外液体，文字位于遮罩之上。基础液体从下方Fly，设2秒，BounceEndIntensity=0.75（现有机制对应界面末段回弹1.5秒）。PowerPoint可将方向11规范读回为3，本机实播确认从下方进入。

两块真正曲线波浪起始位于瓶底以下；后层移动(+3,−8.2)cm，前层(−7.6,−8.5)cm，同一点击组3秒、加速0.3/减速0.7。这些路径参数适配本例画布与构图。波浪宽32cm、波长4cm；初版宽度不足，运动时露出竖边，已扩大水平覆盖范围并重播修正。

最终发布源重新制作后，在本机PowerPoint保存、重开并实播18秒，复看16个分布帧，确认空瓶起态、上升/回弹、双波面斜向上升、结束保持及文字稳定。另以15fps检查270帧，瓶形外4像素抗锯齿容差以外未检出液体颜色。OfficeCLI文件验证通过。静态导出第2页为空瓶是正常的，路径终态需原生播放查看。

这是一次性升起并保持的原生视觉效果，不表达真实容积，不包含无限流体模拟、气泡循环或自动百分比换算；更换瓶形、路径、波长或液位要重新预览全过程。

在仓库根目录运行，输出使用新路径：

```powershell
python examples/flask-wave/build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File examples/flask-wave/finish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Native NEW/native
```

需要Windows桌面PowerPoint、Windows PowerShell及Noto Sans SC/Noto Serif SC。OfficeCLI负责页面与曲线，COM负责真正Subtract、背景填充及动效。
