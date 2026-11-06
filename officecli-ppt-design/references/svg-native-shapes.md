# SVG转真可编辑形状

backend：hybrid。OfficeCLI嵌入caller SVG，PowerPoint对当前自有SVG执行官方原生Convert to Shape命令，输出可编辑custGeom。保留SVG和base；不是截图/PNG替代，也不是把SVG图片改名为shape。

## 来源与机制

大师课9Ewb0Rd3P3E32:06–42:17，35:30在Figma导SVG，35:54插入，36:02–36:09 Convert to Shape后完整填充选项可用。该章转写已读、六个分布画面已看（36:10为转换后原生形状），未称连续逐帧全片。转写34:09误说rounded rectangles，语境实际是rounded hexagons，不能据此声称PowerPoint没有圆角矩形。后续背景窗口与动画另有配方。

官方OfficeDev控制表缓存给出`SVGEdit`（27803，Graphics Tools/Graphic上下文），不是`SmartArtConvertToShapes`（18245）。先激活自有文稿/当前页并选择SVG，GetEnabledMso确认可用，再ExecuteMso('SVGEdit')。本机输入是Type28图形，输出两个Type5自由形状，分别31/7节点；输出可以是多个独立形状，不能强求一定成组。

## 测试输入与边界

调用方提供本地SVG。受测子集：根svg下恰2个直接path、两种不同solid #RRGGBB，各path仅d/fill/fill-rule属性；无text、filter、外部资源。原创例为600×360的Q圆角六边形环＋箭头，反向内轮廓形成真孔。脚本不会下载教程SVG、照片或字体。

`-SourceFill`指定转换后唯一要改色的原始路径颜色，`-EditedFill`指定不同的新颜色，六位RGB不带#。当前验证1B7775→3867A8，第二色E88735保持。几何/渲染验收脚本只针对该环/箭头fixture及20个语义采样；其他SVG需自行定义孔点和轮廓检查，不能拿一次成功覆盖任意SVG功能。

## 从包内复现

构建/结构核验仅Python标准库；像素核验用已有Pillow+NumPy；原生步骤需Windows PowerPoint。新输出拒覆盖；原生UI按队列串行，脚本仅关闭自己的文稿。

```powershell
& $Py "$Skill/scripts/svg_native_build.py" --officecli $OfficeCLI --svg $CallerSvg --out "$Out/base.pptx"
& "$Skill/scripts/svg_native_polish.ps1" -Deck "$Out/base.pptx" -Evidence "$Out/native" -SourceFill 1B7775 -EditedFill 3867A8
& $Py "$Skill/scripts/svg_native_verify.py" --base "$Out/base.pptx" --svg $CallerSvg --evidence "$Out/native" --report "$Out/structure.json" --edited-color 3867A8
& $PixelPy "$Skill/scripts/svg_native_pixels.py" --evidence "$Out/native" --structure "$Out/structure.json" --report "$Out/pixels.json"
& $OfficeCLI validate "$Out/native/edited.pptx" --json
```

base中是实际svgBlip及相同SHA256的SVG media part，并保PNG兼容fallback；转换后目标必须没有p:pic/blipFill，实际含custGeom路径。原生converted.pptx保持两色，edited.pptx为真实单leaf改色，再只读重开核对象/填充保留。两种底层背景带穿过孔区，孔上下应分别露出不同背景色，不能用白形状假孔。

独立核验：caller SVG字节、真custGeom与轮廓数（环2轮廓/箭头1）、非图片leaf、native Type5和Nodes>0；改色前后geometry hash与位置完全相同。像素核20个孔洞/颜色/透明语义点、转前后轮廓IoU、改色保持轮廓、重开全图一致；标题/脚注不受转换污染。

研究结果：4张原生PNG已看，转前后轮廓IoU0.99465，改色后0.99936，重开像素全同；两个路径为真custGeom，孔洞上下透出不同底色，实际leaf改色与保存保留通过，schema零错。SVG/原base未消费。复杂gradient、CSS、mask/clip/filter、字体文字、外链、动画和其他Office版本仍未验证，不能宣传通用无损SVG转换。

包内全新前向已完成：validation/forward-svg-native/。四入口新建CLI底稿→SVGEdit→真leaf改色→保存重开；独立SVG hash/custGeom/轮廓/零图片目标及20语义点通过，转前后IoU=0.99465、改色保持=0.99936，重开全像素相同；4原生PNG已目视，输入SVG/base不变、拒覆盖、schema零错。完整报告在该目录README.md。
