# 同页指定 Pin 开／关标签

backend：`hybrid`。OfficeCLI做基础页与命名形状，PowerPoint COM生成真正InteractiveSequence，原生物理鼠标点击验证。用于地图城市Pin、信息点、注释标签；不是页间链接，也不是普通点击任意空白前进。

## 来源与还原

[One Skill PPT · TOP 10 PowerPoint Animations](https://www.youtube.com/watch?v=o12PIUTskS0)，7:43–9:43。8:13–8:19命名Berlin Label/Pin，8:31–8:47给标签Fade入口并指定Pin触发；8:58–9:02说明只有入口时再点只是重现；9:07–9:17另加Fade出口拖同一trigger组。补采552/556/560秒帧看到出口由主动画移入Pin组；563–568秒放映画面为关→开→关→再开。字幕与实际画面分别核实。

CLI建立一页，6cm大圆按钮`trigger-pin`、独立城市标签`trigger-label`，关闭普通点击/计时翻页。COM步骤必须保持对象身份：

1. `InteractiveSequences.Add(1)`，再`AddTriggerEffect(label,10,4,pin,'',0)`得到入口。Fade=10，OnShapeClick=4，持续0.5秒。保存其真实`Parent`，不要假设原Add返回的空序列仍是有效容器。
2. 直接在该InteractiveSequence调用`AddEffect`在本机报InvalidRequest。通过MainSequence添加标签Fade、`Exit=-1`、持续0.5秒，再`exit.MoveAfter(entry)`，对应原视频“拖进trigger组”。该移动真正迁入同一交互序列；保存后验证只有一个interactiveSeq、无mainSeq、一个entr和一个exit、onClick指向Pin、两效果指向同label。
3. 放映时明确`PointerType=1`（Arrow）并回读，关闭Presenter View；使用正确放映HWND、DPI一致的ClientToScreen坐标与物理鼠标点击。不使用View.Next驱动标签，不使用页切换或宏替代。

## 使用

必须提供全部输出位置；已有输出/目录会拒绝覆盖。Windows本机PowerPoint、OfficeCLI、FFmpeg；Python只用标准库。每个COM脚本单独进程执行，和其它放映/设置UI任务串行。脚本只关闭自己的文稿，不Quit共享PowerPoint。

```powershell
python scripts/pin_toggle_build.py --officecli OFFICECLI --outdir NEW
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/pin_toggle_polish.ps1 -InputDeck NEW/pin-base.pptx -OutputDeck NEW/pin-native.pptx -EvidenceDirectory NEW/polish
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/pin_toggle_test.ps1 -Deck NEW/pin-native.pptx -Report NEW/clicks.json -OutputDirectory NEW/live -Ffmpeg FFMPEG
python scripts/pin_toggle_verify.py --deck NEW/pin-native.pptx --click-report NEW/clicks.json --output NEW/checks.json
```

测试按顺序采图并断言标签区域真实底色：初态关→空白点击仍关→首次Pin开→空白点击仍开→二次Pin关→三次Pin再开。每次目标点击仍为同一页；不把“页码不变”单独视为成功。8秒桌面裁切录屏保留完整鼠标过程，6张状态图和像素断言互相校验。

## 已测范围

研究v1仅保留实际开/关/再开、两次空白负对照及原14项状态/结构证据，后续发现Fade为1ms，不再作为0.5秒渐变节奏通过的成品。可用成品以修后v2为准。源画面和自制作品均确认第三次再开；不据此声称任意快速连击、多个Pin互斥关闭或无限次循环已验证。改形状分组、触发对象名称、时长或层级后重新运行实际点击测试。

精确包内前向`validation/forward-pin-toggle-v2`已通过：新CLI基础页、精修、真实物理点击/桌面录屏、独立15项、最终validate全部过；`live/contact.jpg`9个影片中间态已复看。已有输出目录拒绝覆盖也已实测。失败中间稿不入库。

修正记录：早期`[single].5`被PowerShell解析为null，导致Fade1ms。现显式`[single]0.5`，独立审计新增两Fade滤镜dur=500ms断言。v2新作品、真实三次点击/负对照/桌面影片复看均过，旧记录仅保留状态行为证据，不再作500ms时长证据。
