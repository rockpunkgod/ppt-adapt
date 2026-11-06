# 交互导航：语义、关系与真实命中区域

## 不把四种动作混为一谈

|源行为|真正需要|不能冒充的替代|
|---|---|---|
|普通点击逐项出现|MainSequence中按页面点击推进|指定对象点击触发|
|点击Pin开关标签|InteractiveSequences中指定触发形状|普通onClick|
|点击某页缩放进入、返回|原生Slide Zoom与目标关系|普通slide链接|
|进入一个章节、播到节末后返回|Section Zoom、章节范围与返回关系|单页Slide Zoom|

Hover又不同：大师课实际用MouseOver动作跳状态页，并设置离开卡片的返回区域。点击链接不能叫Hover。

同页Pin开/关已独立完成[指定触发配方](pin-toggle.md)：真正InteractiveSequence、同标签Fade入口/出口、首次开/二次关/三次再开，且空白点击不改变状态。包内新作、14项审计及真实鼠标录屏均通过。它不是普通MainSequence onClick，也不能由其他页的首次入口测试推导。

## 实体透明热点

本轮原生截图发现给文字直接加链接会变主题蓝、加下划线。可把视觉文字和命中区分离，用带真实fill但opacity0的透明形状承载内部链接；热点不带文字，覆盖完整点击区域。

空心圆`fill=none`可能只有轮廓能命中。点圆心失败不能立即判定关系无效；先确认热点填充、实际窗口、遮挡顺序与坐标。

## Zoom缩略图补齐

OfficeCLI能建立原生Slide Zoom结构，但本轮默认poster是1像素占位；PowerPoint不会自动补出目标页画面。生成后虽schema通过，原生截图仍只剩空框。

修复链：用本机PowerPoint导出目标页PNG → OfficeCLI插入隐藏图片carrier → 取得本页对应rId → 同时修改Choice与Fallback的blipFill → 保留原生`pslz:sldZmObj`、目标slideId与返回设置。不要改为普通超链接后声称Zoom修好了。

补poster通过静态检查，不自动代表原生Zoom点击、Return to Zoom、透明背景或Section Zoom均通过。对应测试状态见 [validation.md](validation.md)。

真正分节进入与节末返回已经单独完成新作/实际点击/桌面录屏及包内前向测试，见[Section Zoom配方](section-zoom.md)。它使用真实section GUID与两页节范围，不是把旧SlideZoom配方改名；旧软件Fallback不保留同等行为。

## 双首页防重复进场

大师课01:41:53–02:08:15章中的机制：初始首页有入场动画，然后进入静态主页；各章节统一返回静态主页。将动画主页设成永远循环会导致每次返回重播，与源目标不同。

这一方法的设计关系已蒸馏；完整自动进入、节末返回或Kiosk逻辑仍需按具体作品实测，不凭普通内部链接宣称全部成立。

## 真实点击验收

原生放映测试要记录起点页、目标形状、预期目标与实际页码；非顺序跳转至少覆盖首→末、末→首与跨页。再测试返回与重复进入。

Windows可能打开Presenter View。仅找到`screenClass`不是正确投影窗口的证明；结合`SlideShowWindows`、窗口句柄与画面确认真正显示幻灯片的窗口。如临时改ShowPresenterView等应用设置，测试结束恢复原值。

截图包含演讲者视图控制条、桌面或其他文稿时，不能把其中错误坐标点击当作功能证据。动态失败先区分文件行为、播放器模式与测试定位错误。

小热点另检查DPI：本机144 DPI时同一client可出现1706×1066逻辑尺寸与2560×1600物理尺寸。测试线程可使用per-monitor v2 DPI awareness并与DWM物理边界、GetPhysicalCursorPos交叉核对。大目标碰巧点中不证明坐标单位正确；坐标核对成功仍不代表链接已通过，必须观察实际目标页和返回。
