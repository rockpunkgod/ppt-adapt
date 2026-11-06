# 同数值折线+面积原生组合图

backend: hybrid。来源BV1VTDGYsEmd P16 1623–1998s；全课3秒OCR已读、6分布帧+948/1772s操作补帧核对，不称完整逐秒观看。源最终1812.8s帧确认面积图；1772s组合菜单是中间态。

1. OfficeCLI建真实combo，combotypes=area,line；两系列必须是同一组数值、同一分类、相同主轴。面积负责填色，线表达同一个度量。
2. 面积用蓝→透明渐变并去边线；折线白心圆标记、蓝色3pt线；删除面积的重复legend项，只保留Revenue。
3. 离散月度数据默认smooth=false，避免插值造成虚假峰值；单位写清，value axis从0起。
4. OfficeCLI1.0.160以numLit/strLit写系列，单有chart对象不代表已嵌工作簿。COM ChartData.Activate可物化正确的xlsx，随后把两系列分别绑定B2:B7/D2:D7及A2:A7/C2:C7分类；保存后检查numRef、externalData关系和xlsx实际数值。
5. 改数据时两列必须同步；面积不是第二个度量。运行默认保持输入数字；`-DemoEdit`仅对附带示范数据执行二月29→31的原生编辑测试。

```powershell
python scripts/chart_line_area_build.py --officecli "E:/ppt skills/tools/officecli.exe" --output "E:/work/chart-before-new.pptx"
powershell -NoProfile -File scripts/chart_line_area_polish.ps1 -InputDeck "E:/work/chart-before-new.pptx" -OutputDeck "E:/work/chart-after-new.pptx" -Evidence "E:/work/chart-native-new"
python scripts/chart_line_area_verify.py "E:/work/chart-after-new.pptx" --report "E:/work/chart-audit-new.json"
```

三脚本是六个月示范配方，拒绝覆盖；调用方需要先显式改builder内数据或按这套机制迁移真实数据，不能把示例收入当真实报告。若显式加-DemoEdit，审计相应加--edited。PowerPoint窗口自动化需排队，只关自身presentation和本次ChartData工作簿，不退出共享应用。

已验证：原生PNG、CLI validate零错、1个原生chart含area与line、同一主轴0–60、6值相等、74%→0%alpha、零平滑、真实嵌xlsx与两组numRef。原生工作簿29→31后两条曲线与xlsx/cache一致，最终v9显示31。曾试31→29立即恢复，COM缓存刷新未成功，该恢复实验不计通过；默认成品不依赖恢复。实际审计脚本不容许缺工作簿的早期版本被算通过。

包内三脚本另在`verification/chart-line-area-forward`全新生成默认数据版并通过独立审计、schema及原生画面复看：二月保持29，未启用DemoEdit。此结果与显式编辑后31的研究样稿分别记录，不用恢复失败掩盖数据改动。

研究证据：distill/bilibili-layout-type/course-27/recipe-line-area/，最终after-v9.pptx及native-v9/after.png。此课程其他柱图、表格、饼图、SmartArt能力仍按chapter-notes-p10-18.md待验，不因本组合图通过而整课通过。

封装后独立前向验证：verification/chart-line-area-forward/。直接运行包内三脚本，默认二月29未改、真实xlsx两列与numRef缓存一致、原生PNG已目视、OfficeCLI validate零错误。
