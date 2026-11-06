# 真图表按系列／类别分步出现

backend：**officecli-native**。OfficeCLI直接生成真实clustered-column chart及series/category子元素动画。PowerPoint COM仅只读打开、导图、放映和关闭；本入口不调用ChartData、Series.Formula或Series.Values，不建立xlsx，也不宣称PowerPoint“编辑数据”可用。

## 来源与适用

BV1VTDGYsEmd P21 2181–2322s；2229s菜单整帧已核：整个对象、按系列、按类别、系列内元素、类别内元素。当前只验证**series、category**两项：先讲完整个Alpha产品再Beta用series；按季度比较两产品用category。源使用Wipe，本原创样稿使用Fade隔离验证顺序，不代证源Wipe方向或全部动画预设。

OfficeCLI1.0.160 `help pptx animation --json`明确chartBuild枚举；父路径 `/slide[N]/chart[M]`。参数effect=fade、class=entrance、chartBuild=series或category、trigger=onClick、duration=500（毫秒）。CLI生成2/4条真实chart数据子目标、bldGraphic/bldSub/a:bldChart及graphicEl/a:chart索引。不是整图淡入，也不是矩形假柱。

[AddEffect](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.addeffect)、[MsoAnimateByLevel](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.msoanimatebylevel)、[ConvertToBuildLevel](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.converttobuildlevel)已经Ego Lite查阅；本机COM直接Level/转换/legacy ChartUnitEffect失败，故采用实际支持的CLI入口，不宣称这些COM setter通过。

## 从包内重建

构建和结构核验只用Python标准库；原生验证需Windows PowerPoint、FFmpeg，颜色量测另用已有Pillow+NumPy环境。各路径显式指定、新输出拒覆盖；原生UI按共享队列串行，不Quit任何共享Office应用。

```powershell
& $Py "$Skill/scripts/chart_build_build.py" --officecli $OfficeCLI --out "$Out/chart-build.pptx"
& $Py "$Skill/scripts/chart_build_verify.py" "$Out/chart-build.pptx" --report "$Out/structure.json"
& "$Skill/scripts/chart_build_export.ps1" -Deck "$Out/chart-build.pptx" -Evidence "$Out/native"
& "$Skill/scripts/chart_build_record.ps1" -Deck "$Out/chart-build.pptx" -Evidence "$Out/desktop" -Ffmpeg $Ffmpeg
& $PixelPy "$Skill/scripts/chart_build_measure.py" --native "$Out/native" --desktop "$Out/desktop" --report "$Out/pixels.json"
& $Py "$Skill/scripts/chart_build_verify.py" "$Out/chart-build.pptx" --desktop "$Out/desktop" --pixels "$Out/pixels.json" --report "$Out/checks.json"
& $OfficeCLI validate "$Out/chart-build.pptx" --json
```

两页使用相同原创示范数据：Alpha24/38/46/60，Beta36/29/54/42，Q1–Q4，订单单位，零基线0–70。真实c:chart保留numLit/strLit，可用OfficeCLI改数；这不等于PowerPoint编辑数据工作簿。改系列或类别数量后必须重建对应动画，不假设旧索引会自动随数据增删。

最终OfficeCLI原始文件在本机直接放映时框架从起态可见，只有**2/4次数据点击**：第1页起态无柱→4根Alpha→加入4根Beta；第2页起态无柱→依次Q1/Q2/Q3/Q4双柱。旧xlsx研究文件经过PowerPoint另存规范化后另加了gridLegend框架点击，总3/5次，不能把其计数套入当前CLI入口。重新编辑/另存后必须重新核点击数，不假定规范化没有行为变化。

## 验收边界

独立结构核真实chart关系、两系列原数值和4类别、无externalData工作簿、2/4个真实子目标、按索引顺序、Fade500ms及hold/onClick。实际放映另核8状态物理点击index和8柱显隐；原生PNG复看布局，影片中间帧复看Fade，不只读参数。

record仅接受恰1个可见screenClass窗口并核foreground/自有文稿，使用DWM物理桌面范围，恢复线程DPI状态，仅关闭自身文稿。measure只适用于固定纯色、固定轴和本fixture：从原生PNG定标8柱位置，再归一化16:9桌面测柱内颜色；不是任意照片/复杂图表跟踪器。更换数据范围、颜色或布局后重新校准和放映。

旧xlsx研究与失败保留独立：distill/chart-build/v4曾成功物化两份xlsx；validation/forward-chart-build首次包内第二图Activate失败、选择对照首图值数不符，随后原始chart读取阶段原生chart.dll崩溃，Windows事件标OFFICE_MODULE_VERSION_MISMATCH（POWERPNT16.0.20527.20112、chart.dll16.0.20527.20088）。不把旧成功移作新前向证明，也不据此修改/重启共享Office。需PowerPoint编辑数据时走独立chart-workbook能力，并按当前环境重新验证。当前动画配方没有绕过或伪造工作簿，它明确不包含该能力。

最终包内前向：`validation/forward-chart-build-native-v3/`。从包内重建两页全新PPT，结构/CLI schema、两张原生PNG、6次物理点击/8状态GetClickIndex与8柱显隐全部通过；12秒真实桌面影片的48个0.25秒样本已人工复看，可见按组Fade中间态和稳定终值；readonly前后PPT SHA256完全相同、已有路径构建拒覆盖。原始CLI2/4步与规范化另存稿3/5步区别已写入测试/文案，首轮误用3/5计数的失败保持在native-v2，不拿失败稿代证成功。
