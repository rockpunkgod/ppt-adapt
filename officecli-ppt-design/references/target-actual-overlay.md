# 目标与实际共轴重叠柱

backend：officecli-native。OfficeCLI创建真clustered column图表，两系列共用同一0起点主轴，overlap=100、gapWidth=300；Target先绘制且85%透明，Actual后绘制且不透明。不是把目标和实际相加的stacked chart，也不是两张轴范围不同的图叠放。

来源：27P BV1VTDGYsEmd P15的2472–2697秒。源目标5000不变、实际值变化。已读该段3秒字幕OCR并对照分布整帧；片头“堆叠”说法不作为真实chart grouping依据。P15更早的渐变/图案系列是独立待验范围，本配方不据此声称整章通过。

最小样例两页：Q1–Q4实际45/60/75/90；目标第一页100，第二页仅用CLI改为120，主轴始终0–140。适用于目标与实际同单位且需要看完成程度的场景。若实际超过目标，实际柱应继续向上；本固定样例没有验证超额分支，另换数据复验。

先以chart transparency=85创建两系列，再给`/slide[1]/chart[1]/series[2]`设`color=1C6561`，本机该setter重建第二系列solidFill并清掉alpha。不要假定所有颜色入口都支持`+alpha100`：这一写法在此series setter实验中写成inheritFill。独立审计须实际检查Target alpha15000、Actual无alpha（默认100000），不能只看命令成功。

```powershell
python scripts/target_overlay_build.py --officecli "C:/path/officecli.exe" --out "E:/work/target-new/chart.pptx"
python scripts/target_overlay_verify.py "E:/work/target-new/chart.pptx" --report "E:/work/target-new/structure.json"
powershell -NoProfile -File scripts/target_overlay_export.ps1 -Deck "E:/work/target-new/chart.pptx" -Evidence "E:/work/target-new/native"
python scripts/target_overlay_measure.py "E:/work/target-new/native" --report "E:/work/target-new/pixels.json"
```

构建和结构检查仅标准库；像素检查需Pillow/NumPy。输出拒覆盖；PowerPoint只读打开自有文稿、导PNG、关闭，不访问ChartData/Series.Values/Formula，也不退出共享应用。像素脚本专用于原样1600×900固定颜色fixture，替换内容/颜色需重新定义语义验收。

包内fresh：`validation/forward-target-overlay/`。真chart关系、独立克隆chart part、series顺序/数据/alpha、clustered/100%overlap/300gap及单主轴0–140均过；原生两页已目视。像素从黑色轴独立估计零点和尺度，四实际柱高度对应45/60/75/90；目标100改120后四柱增高，而实际柱区像素、x范围和零点均不变。两系列同x范围证明重叠，不以最终外观冒充stacked数据语义。

此入口是原生字面量chart，CLI可改数值；不含externalData/xlsx，不声称PowerPoint“编辑数据”工作簿能力。它是静态图表方法，没有新增动画待声称。
