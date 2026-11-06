# 字形等距扩张与描边转填充路径

backend：**officecli-native**。Windows GDI+只提取原字形轮廓，Shapely只准备矢量坐标；OfficeCLI创建所有页面、闭合customGeom与填充，没有将结果烘焙成图片，没有COM作者步骤。它不是PowerPoint原生“偏移路径”按钮。

来源：[50P版式设计 P19](https://www.bilibili.com/video/BV1oPHgeiEgb/?p=19)，1010.7–1285.9秒指定文字连续读取，1106/1111/1195秒界面已对照。原软件Illustrator；1111秒实际值是**2毫米、圆角连接**，不是2磅。1195秒为描边转轮廓。后续剪切、渐变、混合不是本条通过范围。

本条只验证 **Arial Black 的6，原字形高11厘米**。脚本拒绝其他文字/字体/高度，避免仍显示固定备份“6”却输出未经验证对象。保留真实文字备份；输出可编辑轮廓不再是可继续输入的活文字。

## 生成

依赖Windows GDI+及已安装Arial Black、OfficeCLI、本机PowerPoint；Python需要NumPy/SciPy/Pillow和`requirements-geometry.txt`中的Shapely2.1.2（实测GEOS3.13.1）。可用`--engine-path`指向独立依赖目录，不自动改变全局环境。

1. `powershell -File scripts/glyph_offset_font.ps1 -Output <new-font.json>`。
2. `python scripts/glyph_offset_geometry.py --source <font.json> --out <new-geometry.json> [--engine-path <deps>]`。
3. `python scripts/glyph_offset_build.py --officecli <exe> --geometry <geometry.json> --out <new-offset.pptx>`。
4. `powershell -File scripts/glyph_offset_export.ps1 -Deck <offset.pptx> -Evidence <new-native-directory>`。
5. `python scripts/glyph_offset_verify.py <offset.pptx> --native <native-directory> --report <new-audit.json>`；具体参数先用`--help`核对。

原轮廓以GDI Flatten(0.1世界单位)变为折线，当前约0.003715厘米，不宣称相对未离散字体Bezier的绝对精度。对折线的2/4/6毫米外扩采用round buffer、quadSegs64；4毫米描边转路径是原边界两侧各2毫米的**闭合填充带**，不是Line属性，也不是均匀放大字号。圆角弧仍为多边形近似。

## 已通过与失败边界

包内脚本从空白生成五页，本机只读原生导出五PNG，文件哈希未变、schema通过。所有结果为Type5原生自由形状、无Line、无图片。独立验证直接从PPT的EMU与路径坐标恢复几何，逐顶点及每条边中点到原轮廓的距离误差小于0.001厘米；不调用Shapely确认自己的结果。

1600×900原生PNG中2/4/6毫米应为9.448/18.896/28.344像素；全部边界偏差≤1.5像素、原字形缺失像素为0。填充描边带在抗锯齿容差外的核心缺失/远处溢出为0；均匀字号缩放负对照被拒绝。没有另存重开测试，不据此声称其结果。

GDI Widen后联合及复合路径旧原型都出现凹角假孔，严格距离检查未过，不能复用。修正使用有效多边形buffer。外扩封闭原本开口会造成合法新孔：本样例孔数1→1→2→2，不把所有拓扑变化都误判为缺陷。当前输入只支持已验证的一外壳及内孔，不扩展到任意字体、多字符、多外壳、自交或字体编辑器功能。
