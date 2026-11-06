# Custom monotone-band glyph envelope

Backend **hybrid**: OfficeCLI creates editable text/target foundation; nativePowerPoint converts actual font glyphs to contours; mathematical sampling maps the contours and OfficeCLIwrites nativecustGeom. It is vector authoring, not an external-render recipe. Output points are editable inPowerPoint, with a separate editable original textbackup.

Source: [BV1oPHgeiEgb P25](https://www.bilibili.com/video/BV1oPHgeiEgb/?p=25) 792–988.3s. Fullrequested captionsegment and exact918/928/931Ego frames were inspected. Illustrator ellipse→Knife partitions→boldtext→targetBringToFront→Object/EnvelopeDistort/MakewithTopObject(Alt+Ctrl+C) is the actual source operation;931shows thetext in the topcap. This recipe is a restricted reconstruction of contour fitting, **not Adobe's live Envelope object or its internal mapping algorithm**, and not a presetWordArt/textWarp substitute.

Domain: a single vertical band described by strictly increasingx samples spanning0..1, with0<=top(x)<bottom(x)<=1. No crossings, overhangs, targetholes, disconnectedpieces, collapsingtips or arbitrary topology. Source's pointed ellipse cap has zero-height tips and is not claimed as replicated. Source/mapped glyphs are closed contourshapes, not livefontcontent; changingtext/target requires rerun, while page1keeps editabletext.

```
python scripts/envelope_band_build.py --officecli EXE --outdir FRESH_BASE
pwsh -File scripts/envelope_band_outline.ps1 -InputDeck FRESH_BASE/envelope-before.pptx -OutputDeck FRESH_OUTLINED.pptx -Evidence FRESH_LOG
python scripts/envelope_band_warp.py --input FRESH_OUTLINED.pptx --officecli EXE --outdir FRESH_WARP
python scripts/envelope_band_verify.py --deck FRESH_WARP/envelope.pptx --model FRESH_WARP/warp-model.json --report FRESH_JSON
pwsh -File scripts/envelope_band_export.ps1 -Deck FRESH_WARP/envelope.pptx -Evidence FRESH_PNG
python scripts/envelope_band_analyze.py --native FRESH_PNG --ffmpeg EXISTING_EXE --report FRESH_JSON
```

Alloutputs fresh; no dependencies/fontinstallation. Fixture uses locally installedArialBoldBOLD with4holes(B2/O1/D1), two original nonpresetbandboundaries. Nativeoutline usesIntersect of thetextwith a coveringrectangle on its own selected slide; original page1textis preserved. Nativegeometry is read from the resultingfile, never from a researchpptx seed. Sourcecommands Move/Line/Quadratic/Cubic/Close are sampled; unsupported commands/formulas reject rather than silently approximate anotherkind of path.

Optional `--targets OWN.json` onwarp accepts exactlytwo targetprofiles for the fixture's two outputpages:

```
[
  {"name":"own-band-a","samples":[[0,0.2,0.8],[0.5,0.1,0.9],[1,0.3,0.85]]},
  {"name":"own-band-b","samples":[[0,0.1,0.7],[0.5,0.3,0.95],[1,0.2,0.8]]}
]
```

Each row isx/top/bottom. Profiles are validated before creating/opening the outputdeck. Mapping:x=u,y=top(u)+(bottom(u)-top(u))*v; Jacobian=bottom-top>0, a nonfolding continuous map before polygon discretization. Targetandglyph are actualcustomDrawingMLpaths. ThelargeXML is sent via supportedCLIbatchJSONpart/xpath/action/xmlfields to avoidWindowscommandline-length limits; no directZIPrewrite authors thedeck.

For anotherword/font, preserve atextbackup and independently revalidate outlinecommand support, spacing, counterholes, bounds and legibility. The fixedBOLDpixelchecker is a diagnostic, not a universalOCR/qualitychecker. Newprofiles/fonts may need finer sampling or designadjustment; positiveJacobian alone does not establish legibility or discretized topology.

Research and packagedfreshforward passed17structure/modelchecks plus6nativePNGchecks and schema validation. Proof: `verification/envelope-band-forward/`; originalbackup, nativeoutline and bothcustomwarps were actually viewed. Sourceglyph197nodes becomes8closedcontours/4121sampledvertices; both customoutputs preserve8contours. MinimumJacobian0.64/0.501586, everyactualoutputvertex within its suppliedband. No rasterpicture orpresettextWarp. Native sourceoutline andbothwarps each have4connectedlettercomponents and4enclosedcounters, and were visually read asBOLD. Countercounts alone do not prove readability.

Untested in the original positive-height recipe: arbitrarytopobject paths, Adobe-liveenvelope editing, zero-heighttips, concaveoverhangs, differenttext/font/targets, exactAdobealgorithm/appearance equivalence. Do not mark the wholearbitrary-envelope familycomplete from this bounded subrecipe. The specific ellipse-cap extension below has its own tests; it does not relax the original custom-target contract.

## 椭圆顶盖：两端收成尖端

适合把短而粗的标题压进分割后的椭圆顶盖。先保留可改文字，再转为真实字形轮廓；上边界沿椭圆，下边界沿浅弯、略倾斜的刀线，两条边界在左右交点合拢。不要用有厚度的圆角条带或预设弧形WordArt代替尖端。处理后重点看两端字形能否辨认、字孔是否还打开；数学上不翻折不等于视觉上可读。

在原制作链的build、warp和verify中加`--ellipse-cap`，outline、export及analyze沿用原脚本：

```powershell
python scripts/envelope_band_build.py --officecli EXE --outdir FRESH_BASE --ellipse-cap
powershell -File scripts/envelope_band_outline.ps1 -InputDeck FRESH_BASE/envelope-before.pptx -OutputDeck FRESH_OUTLINED.pptx -Evidence FRESH_LOG
python scripts/envelope_band_warp.py --input FRESH_OUTLINED.pptx --officecli EXE --outdir FRESH_WARP --ellipse-cap
python scripts/envelope_band_verify.py --deck FRESH_WARP/envelope.pptx --model FRESH_WARP/warp-model.json --report FRESH_JSON --ellipse-cap
powershell -File scripts/envelope_band_export.ps1 -Deck FRESH_WARP/envelope.pptx -Evidence FRESH_PNG
python scripts/envelope_band_analyze.py --native FRESH_PNG --ffmpeg EXISTING_EXE --report FRESH_PIXEL_JSON
```

此分支只生成明确的椭圆顶盖：单位椭圆`(2x-1)^2+(2y-1)^2=1`的上弧与刀线`y=0.14+0.035x+0.08x(1-x)`求两个交点，再归一化并映射字形。左右端高度精确为0，内部高度严格为正；不用小正数垫高尖端。第3页为顶盖，第4页保留原wavy-band作比较，第1页仍有可编辑BOLD备份。默认调用及`--targets`仍要求处处正高度，不能通过输入字段绕过；`--ellipse-cap`与`--targets`不能同时使用。新字体、其他分割位置或更多形状另行制作和预览。

独立核验22项通过，包括实际轮廓的两个端点收敛、椭圆方程、弯曲刀线、内部正高度、八闭合轮廓及边界范围；复用原原生PNG检查，源轮廓、顶盖和正高度比较页都保留四个字母连通分量、四个字孔，并已人工读作BOLD。23秒本机原生播放以5次物理方向键走1→2→3→4→3→4，0额外输入，返回顶盖显示正常；四张原生导图和六个放映画面已看。本例是静态矢量映射，不含动效、Adobe实时包络、任意拓扑或算法等价声明。

[可编辑示例与实际播放](../../examples/ellipse-cap/README.md)。本轮Ego Lite另核该合格来源P25的12张视频元素帧：810秒椭圆、870秒刀线分割、900秒粗体、928秒顶层对象命令、931秒尖端顶盖结果；来源播放量564553。原正高度样稿的17项结构回归仍通过；其自定义目标仍拒绝零高度端点，不用本分支的通过替代原标准。
