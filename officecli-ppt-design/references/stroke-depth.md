# 逐字 / 逐笔画错层与局部渐变补片

backend：**officecli-native**。CLI创建原文文本、每笔闭合轮廓/Bezier、独立闭合补片、局部交叉桥片及真实原生Multiply；PowerPoint COM只用于原生导出、另存重开。源梯度图片是原始黑白场，未把字形或最终混合结果烘焙成整字图片。C10原生Multiply的RGB空间/边界沿用[真混合参考](blend-compositing.md)，不重复该族公式。

源：[50P BV1oPHgeiEgb](https://www.bilibili.com/video/BV1oPHgeiEgb/)，P21 1622.1–1914逐字层序/后字渐变，P25 3158.7–3520.4钢笔局部补片/字内曲边渐变，P48 1762–2019重复局部矩形/钢笔补片。原软件为Illustrator。源ICC/gamma未知，本配方不是作者模板的像素复刻。

## 原子机制

1. **逐字层序**：后字先绘，前字后绘；后字全局明暗场提供退后感，再拉近到有真实交集。单独保留只有字间层序的对照页，避免一条全字渐变冒充所有笔画前后。
2. **接收笔画的局部补片**：三条独立闭合路径分别为竖∩横下区域、竖∩捺边斜带、撇∩靠竖区域。原生Multiply只在该独立补片内发生，每条源渐变方向/范围独立设置。补片不越出实际接收笔画，不以一般投影替代。
3. **局部交叉拓扑**：四笔全局顺序为撇、捺、竖、横；横在竖前、竖在撇前。局部桥片把指定捺∩竖交叉翻到竖前，其余捺不全局置顶。四色诊断页保持相同全局顺序，仅开关此局部桥片，原生同点从绿色竖变橙色捺。
4. **渐变连续与边界**：曲线撇主体为原生cubicBezier；曲边补片是实际几何布尔裁剪路径。桥片保留逻辑交集R∩V，渲染时在R范围内把V左右各扩2个路径单位，避免两个几何边界的抗锯齿重复产生细缝。此余量单列为渲染边界处理，不假称仍是精确R∩V本身。

## 坐标与路径

原创“木木”，每个木四条手绘闭合轮廓，原文另以真实文本框保留。不是任意字体的通用转曲工具，也不把路径图片当作可继续打字的字体。原始600×600路径域在PPT存为600000单位customGeom，原始Bezier控制点/细分轮廓/局部几何/参数均在manifest保留。它们可经脚本或OfficeCLI路径操作修改；不宣称所有PowerPoint UI都提供图片遮罩路径的Edit Points按钮。

凸裁剪用Sutherland–Hodgman。曲线主体保留native cubicBezier，曲边补片裁剪使用80段参数细分得到的多边形近似，明确不是精确解析Bezier布尔求交。当前仅验证原创少量交叠笔画，不覆盖任意字体、复杂复合孔/自交、自动拆笔或完整汉字拓扑。

每条Multiply第二源必须使用**其自身可见几何bounds**采样，不能用整个字形画布坐标直接拉伸。probe-v1实测整字坐标让竖笔补片缩成细条；probe-v2按笔画bounds修正才真正全宽对齐。含小数bounds应保留精度，当前在每路径单位2倍分辨率、像素中心采样原创黑白源场，主色和乘法由PowerPoint原生执行。

## 对照和验收

包内5脚本：`stroke_depth_geometry.py`（被以下脚本导入）、`stroke_depth_assets.py`、`stroke_depth_build.py`、`stroke_depth_native.ps1`、`stroke_depth_verify.py`。资产生成需要Pillow，验证需要NumPy/Pillow；CLI builder只用标准库，禁止用python-pptx替代。独立输出目录/文件显式给出，默认拒绝覆盖。

```powershell
python scripts/stroke_depth_assets.py --outdir NEW/assets
python scripts/stroke_depth_build.py --officecli OFFICECLI_EXE --assets NEW/assets --outdir NEW/deck
officecli validate NEW/deck/stroke-depth.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/stroke_depth_native.ps1 -Deck NEW/deck/stroke-depth.pptx -EvidenceDirectory NEW/native
python scripts/stroke_depth_verify.py --build NEW/deck --assets NEW/assets --native NEW/native --report NEW/verification.json
```

不要把独立补片源图换成全字混合后的PNG。保存manifest的原文、控制点、逻辑交集、渲染余量、补片接收者、sourceFields的fractional bounds与输入SHA。改路径或画幅时应重算几何和源坐标，再做原生核验。

作品三页：①只有字间层序；②局部笔画错层完整原创双木；③四色全局层序不变、局部桥片开/关。

至少检查：三页原文和原生闭合路径/Bezier/Multiply节点保存重开保留；输入SHA不变；原生两轮图像一致；真实几何/层序/源场独立像素重建；笔画外3px边界以外全像素背景不变；逻辑桥片、渲染桥片及三个独立补片各自限于正确接收笔画；桥片到原捺接缝不出现暗线；渐变白端连续且往遮挡处变暗。

历史阴性须保留：full-v1虽画面成立但补片仅RGBA场，严格中心RGB43/45；full-v2改独立真实路径后64/66，四色陡梯度偏差最大6及局部接缝46级，未入库。后续版本调整原创渐变宽度和桥片接缝余量，未把验收RGB阈3放宽。

完整研究v3：`full-v3/verification.json` **71/71**；三页原生导出/保存重开逐像素0差，输入SHA不变。每页独立几何/层序/梯度样点10378/10338/10406个，RGB最大0/1/2（验收阈3未放宽）；笔画轮廓3px以外另871846/871846/877118个全像素背景0差。桥片同点从(57,219,169)变(255,161,66)，接缝7点均(255,161,66)、最大0差；渐变白端245→244连续，逐步暗到101。80/160段曲线细分最大差0.00657路径单位，是收敛对照而非解析误差上界。

独立包内fresh：`validation/forward-stroke-depth/`，由上述包内5脚本全新生成资产与三页PPT，原稿/另存roundtrip均CLI validate通过。原生3图→另存→重开3图，`verification.json` **71/71**，每页样点10378/10338/10406、RGB最大0/1/2且阈3未放宽；约87万/页笔画外3px之外全像素背景0差；桥片接缝0差、白端连续/暗化、真实独立补片/Bezier/原文保留、输入SHA不变、重开逐像素0差。三页已人工查看。均为静态原生证据，不称桌面录屏或动态验证。阴性仍只在研究目录保留，不混入fresh作品。
