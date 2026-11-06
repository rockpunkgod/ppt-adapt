# 同心饼图/外环同值与百分比卡片

backend：officecli-native。真实pie与doughnut共享同一类别/顺序/数值/起始角，图表区和绘图区noFill，外环内孔大于内饼外径，中间透出真实背景。外环重复相同比例作装饰，不代表第二个度量。

来源BV1VTDGYsEmd P17 1680–2964秒，全部可用3秒OCR已读，1680/1830/2019/2220/2745/2960整帧已看。前半由两张图同心叠放，后半拆成三张v/100-v卡片。手绘三角装饰和照片背景不包含。2745秒还是15/20与旧80标签的中间态，不能照抄为正确结果。

默认第一页两图55/20/15/10；CLI clone后真实set两图series values为40/30/20/10，同时更新四标签。第三页明确是独立的65/25/10示例，各环存v和100-v；绿色代表卡片命名份额、蓝色代表余数，图内说明颜色角色，不误当与前页四分类颜色一一对应。

```powershell
python scripts/concentric_chart_build.py --officecli OFFICECLI --out NEW/chart.pptx
powershell -NoProfile -File scripts/concentric_chart_export.ps1 -Deck NEW/chart.pptx -Evidence NEW/native
python scripts/concentric_chart_verify.py NEW/chart.pptx --native NEW/native --report NEW/audit.json
```

builder仅标准库；验证需Pillow/NumPy，导图需本机PowerPoint。输出拒覆盖、只关自有文稿，不Quit共享应用，不访问ChartData/Values/Formula。`--values`、`--changed-values`、`--cards`分别接4/4/3个有限非负百分比，各组总和100；校验器仅验默认固定数据/1600×900/主题夹具，换值另核实际输入/扇角/标签。

两图theme accent1–4及varyColors=true，原生实测同索引同色；三卡用accent1/2。改theme/style后须重新核颜色，不能只改图例。firstSliceAngle=0在本机为12点起顺时针，外环内径85%、卡片68%；不外推全部主题/Office版本的自动布局。

包内最终fresh：`validation/forward-concentric-chart-v2/`。7真chart独立部件，全部数据/标签/总和/图心/起角/内径通过独立OPC；第二页实际改值不污染第一页。三原生页已目视；720角度采样，内饼/外环比例误差<0.6个百分点，有效角度颜色对应100%，空隙上下透出不同背景色，三卡角度与65/25/10相符。中心数字的位置已修正。

源风险70/20/10/5（总105）的独立输入测试在创建PPT前拒绝，未生成文件，见`validation/concentric-invalid-total.json`。三卡是一个整体分拆，不适用分母各异的独立指标。此入口是原生字面量chart，CLI可改值，不含externalData/xlsx，不声称PowerPoint“编辑数据”工作簿能力；它是静态配方。
