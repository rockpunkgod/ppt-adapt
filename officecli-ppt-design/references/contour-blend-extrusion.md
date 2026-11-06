# 轮廓连续挤出、长影、等角分层

来源：50P P25 187.4–262.3、602.4–746、2815.5–3120.8；P26 3076.4–3200.5。后台：**hybrid**（OfficeCLI负责填充矢量与排版，PowerPoint原生Union/Intersect处理长影）；不含Multiply渐变合成/真3D模型。

适用：同一轮廓的两个对应端面，沿位移/尺度/颜色生成连续挤出；将字形沿45°位移的扫掠合并裁剪；将描边轮廓化为填充带，做等角顶视平面与稀密切面。当前闭环目标仅自产P/Arial Black/5cm高/0.3mm半描边宽默认参数；其他字形、宽度、剪裁需重新过oracle与原生图。

决定规则：

1. 有对应节点才逐节点插值，不把两种拓扑不兼容字形硬配对。200中间步是202个端面总数；另加正面盖板以独立定色。
2. 连续平移扫掠可以由起末完整字形与每条边扫过的四边形组成，包含孔洞边界，再真Union/Intersect。必须拿完整字形多位置采样作独立轮廓oracle，不能因对象减少就称等效。
3. 切面轮廓应为**填充带**，不可用细线条的原生描边代证。先验证真实等宽几何，再用等角顶视仿射矩阵`[[sqrt(3)/2,-sqrt(3)/2],[0.5,0.5]]`投影；垂直偏移叠202/32片，分别对应源200/30中间步。
4. 顺序由远到近，近/远端颜色插值；孔洞保留为反向子路径。长影扫掠可能自然填掉初始孔洞，不能强制给最后长影打同样一个孔，必须以oracle为准。
5. 原生merge先存primary，再存Range，最后调用MergeShapes；不要在PowerShell COM链式调用里直接合并。无论失败/成功只关闭本脚本自有稿，不Quit共享PowerPoint。

输入要求：不需作者照片/模板/外站。只读取本机字体或调用方自行提供合法轮廓；未捆字体。以OfficeCLI从零创建基础页和全部custGeom。几何预处理不是external-render，最终核心是真原生可编辑freeform。

验收：真实Type5/Nodes可读与保存重开；436片节点坐标独立公式检查；202完整字形union与边界mesh几何；原生长影像素/裁剪；band与原边界最近距离独立检查；逐页原生PNG。旧GDI Widen缺口、最初COM链式Range失败全部保留，不入通过结果。

状态：包内fresh已通过，见 validation/forward-contour-blend/README.md。旧Widen切面只作失败证据。静态PNG/真实可编辑节点已验；没有动画实播声明。

## 包内复建步骤

依赖：Windows本机Arial Black、PowerPoint、OfficeCLI1.0.160；基础Python；独立像素检查Python需Pillow/NumPy；几何预处理/薄带检查需固定Shapely2.1.2。脚本通过`--geometry-deps`显式接受独立依赖目录；本机可使用`E:/ppt skills/tools/geometry-deps`。不修改全局环境/不自动联网安装。若不可用，停止对应几何预处理，不能退回已失败Widen带。

设定`$skillRoot`为本技能目录，`$taskPython`为基础Python，`$taskPixelPython`为含Pillow/NumPy的Python，`$taskCli`为OfficeCLI，`$taskGeometryDeps`为Shapely路径，`$taskOut`为**不存在**的新目录。新建taskOut之后输入/中间文件名均需不存在。

```powershell
if(Test-Path -LiteralPath $taskOut){throw 'Refusing overwrite'}
New-Item -ItemType Directory -Path $taskOut | Out-Null
& "$skillRoot/scripts/contour_blend_outline.ps1" -Output "$taskOut/glyph.json"
& $taskPixelPython "$skillRoot/scripts/contour_blend_band.py" --input "$taskOut/glyph.json" --output "$taskOut/band.json" --geometry-deps $taskGeometryDeps
& $taskPython "$skillRoot/scripts/contour_blend_build.py" --geometry "$taskOut/band.json" --officecli $taskCli --out "$taskOut/work"
& $taskPixelPython "$skillRoot/scripts/contour_blend_sweep.py" --model "$taskOut/work/model.json" --out "$taskOut/sweep-check"
& $taskPixelPython "$skillRoot/scripts/contour_blend_band_verify.py" --model "$taskOut/work/model.json" --out "$taskOut/band-check.json" --geometry-deps $taskGeometryDeps
& "$skillRoot/scripts/contour_blend_polish.ps1" -Deck "$taskOut/work/blend-base.pptx" -Output "$taskOut/blend-native.pptx" -Evidence "$taskOut/native"
& $taskPython "$skillRoot/scripts/contour_blend_geometry.py" --deck "$taskOut/blend-native.pptx" --model "$taskOut/work/model.json" --out "$taskOut/geometry-check.json"
& $taskPixelPython "$skillRoot/scripts/contour_blend_verify.py" --deck "$taskOut/blend-native.pptx" --model "$taskOut/work/model.json" --native "$taskOut/native" --out "$taskOut/native-check.json"
& $taskCli validate "$taskOut/blend-native.pptx" --json
```

查看native的三张PNG，逐项检查盖板孔洞/连续挤出/裁剪位置/稀密片层。`native-objects.json`保存前后Type5/Nodes证据。所有输出拒覆盖；PowerPoint只关闭自身打开的稿件，不退出共享应用。静态技巧不宣称已做动画实播。

默认工程：三页，548个构造时custGeom；真合并后442个custGeom（其中436个插值切面），3页形状总数211/8/239（含文字）。正面孔洞真实透明，长影孔洞按扫掠自然变化。最终等角轮廓带有4环/844节点，源200与30中间步对应202与32片。更改默认形态/参数需重做独立oracle与native检查。


包内独立前向：从本机P字形重新提轮廓，没有复用研究成品。202完整位置参考与106边界扫掠IoU=.9991345、内部漏像素0；buffer薄带与原边界最近距离网格IoU=1.0；436片坐标/颜色逐项核对，最大坐标差0.0000451cm；真实原生长影像素IoU=.988239，近远孔洞背景像素误差0。106Union→1个247节点freeform，裁剪后141节点；等角面844节点，保存重开Type5不变。三PNG逐页复看、OfficeCLI schema与4项拒覆盖、基础稿SHA不变均通过。
