# 为字面量图表补齐真实工作簿

backend=`hybrid`。OfficeCLI基础chart可能只有numLit/strLit：图形是原生对象，但不能因此保证PowerPoint“编辑数据”已经有真实xlsx。交付需要在PowerPoint改数据时，补工作簿、显式绑定单元格，并核xlsx、公式、缓存和原生画面。

本工具对已有的**共享类别、完整连续缓存、尚无工作簿**图表生成绑定方案。已实际通过visual-design单系列横向条形和storytelling双系列簇状柱形；没有把这两例外推为散点、气泡、稀疏点、多层分类、外部链接或所有图表类型通过。已有工作簿不走此补齐工具，避免清理用户既有表格。

当前强制限每文稿一张chart，Python方案与COM入口都在打开Office前拒绝多图表。旧单图表正向证据保持；后续两图表同稿的独立前向在工作簿激活/数据提交阶段失败，进一步诊断发生chart.dll崩溃，WER明确记录OFFICE_MODULE_VERSION_MISMATCH。没有把研究中偶次双图成功当可靠能力，不自动拆改用户文稿绕过此边界。图表分步动画本身与补xlsx是独立能力；不要为了动画验收去触碰未要求的数据编辑接口。本轮错误报告在distill/chart-build/first-chart-stages/windows-error-events.json；未执行Office修复/升级。

## 默认保留原数据

先保证OfficeCLI对基础PPT做过close/flush，再运行：

```text
python scripts/chart_workbook.py plan ./input.pptx --out ./new/plan.json
powershell -NoProfile -File scripts/chart_workbook_polish.ps1 -InputDeck ./input.pptx -OutputDeck ./new/editable.pptx -Plan ./new/plan.json -Evidence ./new/native
python scripts/chart_workbook.py verify ./input.pptx ./new/editable.pptx --report ./new/structure.json
```

输出PPT、方案、审计报告和证据目录均选新路径并拒覆盖；输入由路径和SHA256锁定。COM仅关闭自己的presentation和本次ChartData工作簿，不Quit共享应用。

方案读取实际slide/shape/chart关系、系列名、类别和值，支持OPC相对/绝对Target。COM通过ChartData.Activate取得真实工作簿，将共同类别写A列、各系列名称写B1/C1…、值写其下，再将每个Series.Name/XValues/Values明确绑定到真实工作表区域；不依赖Activate自动猜对数据绑定。类别按文本写入，数值用double。

默认不改业务数值，也不重设chart type、fill和axis range。脚本比较COM原生系列值和前后格式快照，导出before/after PNG。独立验证还会核嵌入xlsx关系与所有类别、系列名、值，numRef/strRef公式和cache，检查页内标题/说明/人工标注未丢失。最后看原生PNG，不能只添加一个空xlsx关系。

## 显式改值测试单独留副本

需要验证编辑行为时，提供独立JSON，例如：

```json
[{"slide":1,"shape":"survey-chart","series":1,"point":2,"value":35}]
```

再次对原始输入运行polish，选新的OutputDeck/Evidence，并加`-EditsPath ./edits.json`；verify加`--edits ./edits.json`。工具写真实Cell.Value2并Refresh，要求原生Series.Values反映修改。序号从1起，shape名从方案读取；不默认修改任意用户数字。测试值应与其他结论/标注保持一致，若改了结论依据，还需同步重算说明。

默认成品与编辑测试副本分开保存。不能从一次成功改值推断“立即恢复”也通过；如要求往返恢复，需要另测。

## 已测结果

- visual-design默认保留25/33/63/70/84，5类别与受访者占比系列正确；独立测试将线上购买33→35。84%焦点不变。
- storytelling默认保留北区80/94/115/138、其他地区82/86/88/91；独立测试将北区2月94→96。1月→4月+73%标注仍对应原端点。
- 四份PPT的xlsx、绑定公式、缓存、原生值和schema均通过；类型、颜色、主轴范围保持。默认两个before/after原生PNG各自整页逐像素相同，改值副本条柱高度/数值标签实际更新。两个原始输入SHA256仍保持。

研究与结果位于validation/chart-workbook-repair。该机制复用chart-line-area中已验证的Activate+真实区域绑定经验，补齐现有visual/story配方；不改变其信息结构和视觉内容。独立封装入口前向另见validation/forward-chart-workbook。
