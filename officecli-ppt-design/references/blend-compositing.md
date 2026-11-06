# 真混合与局部明暗：Multiply / Overlay / SoftLight

源：50P课程 [BV1oPHgeiEgb](https://www.bilibili.com/video/BV1oPHgeiEgb/)，P22 2089.3–2184.8、P23 1554.7–1640.6、P25 3158.7–3520.4、P48 1762–2019。源为Illustrator，不是PowerPoint；RGB/GPU预览已见，但源ICC与混合gamma未知。因此本配方复现明确定义的混合机制，不声称逐像素等于作者未知颜色设置。

## 后端和可迁移范围

|原子|backend|已测对象/范围|
|---|---|---|
|不透明源的正片叠底|officecli-native|单图片对象 `a:blip/a:fillOverlay blend="mult"`，额外图片填充。|
|不透明源的叠加|officecli-native|同结构 `blend="over"`；48个RGB对确认是W3C Overlay而非普通覆盖。|
|Multiply后50%|officecli-native|`fillOverlay` 内第二图片的 `alphaModFix amt="50000"`；与普通50%叠图分别实测。|
|局部渐变Multiply|officecli-native|同原生效果，第二图为原创黑白渐变RGBA，alpha只在多边形内为1；无预烘焙结果替换该原生图。|
|SoftLight、局部SoftLight|external-render|明确W3C公式、sRGB编码值、半透明覆盖率运算，确定性生成PNG，再由CLI嵌入。|
|局部Overlay|external-render|当前保守配方采用显式蒙版外部合成；不外推半透明原生 `over` 等价性。|
|Screen（辅助控制）|officecli-native|`blend="screen"`，48个RGB对核验。|

OfficeCLI负责所有页面/图片/文字和真实DrawingML效果；COM仅原生导出、SaveAs重开验证，没有COM作者操作。不透明Overlay不得未经测试扩展到半透明源：本机原型把内层alpha设50%的 `over` 并不等于通常Overlay后50%线性插值，已排除该配方。SoftLight的外部实现是当前已验证办法，不代表穷尽所有Office effectDag的可能。

## 公式和输入

背景B、源S均归一化0–1；本配方默认直接在sRGB编码值上运算，coverage alpha不做gamma变换，背景不透明：`C=(1-alpha)*B+alpha*Blend(B,S)`。

- Multiply：`B*S`。白色源不改变背景、黑色源压黑。先Multiply再50%不等于Normal50%。
- Overlay：B≤0.5时 `2*B*S`，否则 `1-2*(1-B)*(1-S)`。
- W3C SoftLight：S≤0.5时 `B-(1-2*S)*B*(1-B)`，否则 `B+(2*S-1)*(D(B)-B)`；B≤0.25时D为`((16*B-12)*B+4)*B`，其余为`sqrt(B)`。

生成器同时保留“解码sRGB→在线性光计算→编码回sRGB”参考，不能把两种空间混称一种。最终量化裁剪0–1后 `floor(255*C+0.5)`。PNG附sRGB ICC，参数、尺寸、RGB调色板、蒙版、文件SHA256在manifest中。输入仅原创数值色场，不捆绑作者照片或教程素材。

## 制作步骤

包内脚本：`blend_compositing_assets.py`、`blend_compositing_build.py`、`blend_compositing_native.ps1`、`blend_compositing_math_audit.py`、`blend_compositing_measure.py`、`blend_compositing_verify.py`。资产/测量需要NumPy、Pillow/ImageCms；CLI builder及最终结构验证仅标准库。不要用研究旧文件名运行。所有输出参数显式填写，已有目录/文件拒绝覆盖。

1. 用资产生成器生成背景/源/局部RGBA/外部SoftLight等文件到新目录。原生效果的源图和背景图分别保留；外部结果也保留原始输入、公式、蒙版和输出，不能宣传为可编辑原生SoftLight。
2. builder以CLI添加底图，添加离页隐藏源缓存得到同页rId。原生混合XML以原图 `a:blip` 的fillOverlay节点引用缓存源rId，第二填充使用stretch/fillRect。不要误以第二张Normal图片覆盖来代替。
3. 标明external的面板直接嵌入确定性PNG；文字标题仍可编辑。黑白渐变只在局部alpha范围生效，蒙版外必须和原B一致。
4. CLI validate；用本机PowerPoint导出、保存到新文件、关闭重开再导图。验证所有原生节点保留、重开像素一致、RGB见证和局部边界，同时保留未通过半透明Overlay诊断。

调用顺序（路径替换为实际新目录，`python`需有以上资产依赖）：

```powershell
python scripts/blend_compositing_assets.py --outdir NEW/assets
python scripts/blend_compositing_build.py --officecli OFFICECLI_EXE --fields NEW/assets --outdir NEW/deck
officecli validate NEW/deck/blend-compositing.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/blend_compositing_native.ps1 -Deck NEW/deck/blend-compositing.pptx -EvidenceDirectory NEW/native
python scripts/blend_compositing_math_audit.py --fields NEW/assets --report NEW/math-audit.json
python scripts/blend_compositing_measure.py --probe NEW/deck --fields NEW/assets --evidence NEW/native --deck-name blend-compositing.pptx --report NEW/pixels.json
python scripts/blend_compositing_verify.py --deck NEW/deck/blend-compositing.pptx --measurements NEW/pixels.json --math-audit NEW/math-audit.json --fields NEW/assets --report NEW/verification.json
```

输出标题/面板及manifest中的backend标签必须保留。源图片缓存是同页离画布隐藏图片，效果rId只能引用本页关系，不得在移动面板到另一页后保留旧rId。

原生图片内部效果不是Illustrator任意两矢量图层的全局BlendMode。此处局部多边形机制并不替代汉字转曲、拓扑调整、组隔离/挖空或完整立体字设计。更换图片、透明度、编码空间、软件版本或局部边界应重新量化，不能直接把本机2级以内容差当通用跨软件保证。

## 研究与封装前向证据

`native-probe-v1` 两页SDK通过，本机mult最大1级、over/screen最大2级、Multiply50最大2级；局部Multiply858内/外样点最大1级；保存重开两页逐像素相同且全部原生节点保留。独立公式/范围控制14项通过。`over50` 阴性保留，不入库该组合。研究帧和日志不是包内fresh前向的替代。

独立包内fresh：`validation/forward-blend-compositing/`。从上述包内脚本全新生成RGB资产及两页PPT，CLI validate通过，本机导出2图→另存→重开2图；`verification.json` **71/71**，输入SHA256不变，保存重开两页逐像素0差。每面板48个RGB见证，原生Multiply最大1、Overlay/Screen/Multiply50最大2级；外部SoftLight和局部SoftLight/Overlay原生显示最大1级。局部原生Multiply另858内外样点最大1级。两页均已人工查看，无标签截断；这些是静态原生像素证据，不称桌面录屏或动态验证。失败over50仍只在独立研究probe保留，fresh没有该组合。

官方依据：[微软FillOverlay](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.filloverlay?view=openxml-3.0.1)、[BlendModeValues](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blendmodevalues?view=openxml-3.0.1)、[W3C混合公式](https://www.w3.org/TR/compositing-1/)。
