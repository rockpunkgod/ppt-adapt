# 单文本框逐字 Wave

backend: **hybrid**。OfficeCLI制作活文字，PowerPoint创建真正Wave强调动画；保留逐字迭代、往返路径和四段旋转，只调整所需幅度、时长与循环。

## 制作

对单个文本框调用原生`AddEffect(82)`及`ConvertToTextUnitEffect(effect,1)`。本机生成preset34、`iterate type=lt`、`tmPct=10000`，即By Letter/延迟10%；不要拆成六个文本框模拟上下平移。

原生Wave包含一个竖向往返Motion和四段+25°、−25°、−25°、+25°旋转。竖向Motion保留自动反转和平滑开始/结束各50%；改红柄幅度对应的`MotionEffect.Path`终点，不丢旋转或改变整框位置。本例−0.04与−0.08均为相对页面高度，绝非厘米。`MotionEffect.ByY`回读100不等于本例路径终点，不能拿它当实际幅度。

时长先由`Timing.Duration`统一缩放，再将UntilEnd的`repeatCount=indefinite`写到**Wave父节点**。不要给各子行为各自重复，也不猜COM的特殊RepeatCount数值。脚本以OfficeCLI精确设置该属性、close后再原生重开保存，保留原生波浪机制。

整词周期还包含字母错峰。本例6字母、10%延迟，0.5秒设置的整词周期约0.75秒，2秒设置约3秒；不能把动画面板Duration直接拿去同步另一个整词循环。不同字数与文本结构应重新预览，不外推多行、单词拆分或中文字符计数。

## 运行

```powershell
python scripts/text_wave_build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -File scripts/text_wave_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/text-wave.pptx -Native NEW/native -Officecli OFFICECLI
```

需要Windows桌面PowerPoint、Windows PowerShell、OfficeCLI、Python和示例字体Noto Serif SC/Noto Sans SC；字母为Arial。使用新输出，不覆盖既有时间轴。四页分别为0.5秒、2秒同幅、2秒双倍路径幅度及静态对照。24pt tracking用于本例的疏朗构图和分字观察，不是Wave所必需。

## 实播范围

`examples/text-wave/`保存可编辑PPTX与44秒原生播放。六次真实鼠标/方向键动作覆盖三页独立启动和翻至静态页；20个分布帧已复看。三段影片的六字母均持续运动，低幅质心纵向范围约32px，高幅约60px，静态对照最大变化不足0.06px。路径幅度翻倍不意味着每种字形的像素质心范围精确翻倍，因为还保留了逐字旋转。

慢放相邻字母的实测滞后约0.167–0.233秒，吻合0.2秒设置与30fps采样精度；快放约0.033–0.067秒，吻合0.05秒。慢放重复相关峰约3秒，两个幅度均超过0.998；快放峰在0.733/0.767秒两侧，吻合0.75秒。每段单独检查，不用快放代证幅度或用单轮代证循环。

保存重开后的文件确认三个单活文本框、原生preset34、letter10%、父循环、一往返Motion及四段旋转，静态页无动画；OfficeCLI验证零错。只声明该单行六字母示例和列明参数的实际结果；返回重启、任意文本、多行、多对象组合及UntilNextClick均未由此通过。

制作依据：[完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)4:41:48–4:42:45，另有3:59:40–4:00:20。复读完整对应归档转写、核看16933/16940/16948秒归档操作帧：0.5秒参数可见；Repeat菜单停在UntilNextClick只是中间态，最终UntilEnd依据完整口述及本机实播，不偷换菜单中间状态。本轮没有新增在线播放覆盖。
