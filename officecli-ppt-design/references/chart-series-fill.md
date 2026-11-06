# 图表系列渐变与原生横纹填充

来源：BV1VTDGYsEmd P15，2103–2325秒渐变、2334–2469秒图案。完整可用3秒OCR已读，2190/2320/2450秒整帧已看：渐变作用在真正数据系列，图案由前景荧光绿与不透明暗背景组成。源明确提示图案背景应接近整页暗底，以减少色差，不把暗条误称透明窗口。未精确辨认源横纹预设枚举，样例选窄横纹作为同机制变体。

两条后端分别归类：渐变为officecli-native，图案为hybrid。渐变用OfficeCLI chart `gradient=28D9E7-102338:90`，真实c:ser/a:gradFill；横纹由PowerPoint `Chart.SeriesCollection(1).Format.Fill.Patterned(30)`，即msoPatternNarrowHorizontal，写成pattFill/narHorz，再设ForeColor/BackColor。使用原生格式API，不访问ChartData、Series.Values或Formula，不把普通shape条纹盖在柱子上。

样例保留同一45/60/75/90数据及0–100主轴；第二页只是复制图表后换图案，数据未改变。两页都是普通原生柱图，不使用三维透视或非零轴制造差距。渐变逐根柱从顶端青色到底端深蓝，按每根柱自身高度归一；视觉渐变不代表第二个数据指标。本例两端均不透明，未声称复现源背景图片或可调透明色标；图案暗条同样不透明。

```powershell
python scripts/chart_series_fill_build.py --officecli "C:/path/officecli.exe" --out "E:/work/fills-new/base.pptx"
powershell -NoProfile -File scripts/chart_series_fill_polish.ps1 -Deck "E:/work/fills-new/base.pptx" -Evidence "E:/work/fills-new/native"
python scripts/chart_series_fill_verify.py "E:/work/fills-new/native" --report "E:/work/fills-new/audit.json"
```

构建仅标准库。原生精修需要已装PowerPoint，验证需Pillow/NumPy。输出拒覆盖，只关闭自有文稿，不退出共享PowerPoint；保留CLI base不变。原生创建新图案页、保存styled.pptx并只读重开两页，分别导图供对照。任意数据/颜色的制作可修改样例参数，但像素脚本只针对固定1600×900诊断fixture，不能直接外推。

验收须检查真实chart关系/series数据/主轴及实际gradFill/色标/angle、pattFill/fg/bg。原生像素需证明每根柱独立渐变归一、横纹在柱内交替且不延伸出柱、两页数据对应高度保持；保存重开全图应一致。图案第一根可见亮线不一定等于柱顶，窄横纹的暗条会产生一个纹理周期的偏移；测量应由可见条纹推导周期并单独核数据，不能把亮线顶部当精确数据边界。

包内fresh前向已通过：`validation/forward-chart-series-fill/`从三个包内入口重新创建base、精修两页、保存重开与独立对象/数据/像素审计，schema零错，原生两页已目视且重开全像素相同。首亮线高度测量的两个原型失败及纹理周期修正保留在`validation/chart-series-fill-measure-corrections.md`，没有靠改图表或数据制造通过。本配方不宣称原生编辑工作簿或任意chart类型、纹理缩放/印刷无摩尔纹。

官方接口已用Ego核查：[ChartFormat.Fill](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartformat.fill)、[FillFormat.Patterned](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.fillformat.patterned)。本地PIA枚举读值与原生回读Pattern=30/FillType=2一致。
