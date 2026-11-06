# FIELD / STUDIO：双首页导航

`double-home.pptx`是一套8页可编辑的交互演示：先自动呈现两个章节入口，然后停在静态首页。点击01或02进入对应章节；每章两页正文，播完回到静态首页，入场不重复。以“观察证据”和“建立设计语言”为两条内容线，采用薄荷底、深绿/陶土色章节标识和一致的字体层级。

页面顺序：动画首页1 → 静态首页2；A入口3 → 正文4 → 正文5 → 首页2；B入口6 → 正文7 → 正文8 → 首页2。入口页自动推进，正文由点击推进。不要将放映改成指定幻灯片范围；本例使用所有幻灯片及原有计时。

`playback.mp4`拼接了最终文件的两段本机PowerPoint录制：首次播放4秒、两轮章节往返54秒。前段用真实F5启动，无幻灯片点击即自动到静态首页；后段12次物理点击，四次节末返回都不重播。`preview.jpg`保留正文、缩放中间态和返回画面。没有将单项Section Zoom的通过结果当作本组合的验收。

重建时先执行现有[Section Zoom制作链](../../officecli-ppt-design/references/section-zoom.md)的build、prepare、convert，再运行：

```powershell
powershell -NoProfile -File officecli-ppt-design/scripts/double_home_compose.ps1 -InputDeck NEW/section-final.pptx -OutputDeck NEW/double-home.pptx
```

只接受自身的五页新底稿和新输出路径，不覆盖业务稿。需要本机OfficeCLI、Python、Windows PowerShell与桌面PowerPoint，示例字体Microsoft YaHei；无需下载素材。新增部分仅编排首页状态和已有原生效果，没有另建演示引擎。悬停、Kiosk、任意快速点击或不同节结构不在当前通过范围内。
