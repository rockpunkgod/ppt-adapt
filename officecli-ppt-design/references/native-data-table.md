# 原生数据表：分组表头、真实合并与必要边框

backend=`officecli-native`。数据、合并、填色和单格边框全部由OfficeCLI生成；PowerPoint COM仅以只读方式核验真实表格和导PNG，不改数据。适用于PPT中的小型数据比较表，不能用文本框/矩形模拟可编辑单元格。

来源[BV1VTDGYsEmd P14](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=14)：867–1044秒原生合并/拆分、1074–1308秒数字表层级与单位、2211–2754秒去冗余线保独立列。已读相关字幕OCR、看885秒合并按钮及2388秒“去中线别合并整行”操作帧；这是采样研究，不称全片逐秒口述转写。

## 设计判断

- 标题说明结论，单位与统计口径明确，行标题/列标题/数据/合计各有层级。数值统一精度并右对齐，正文行不要为了装饰随意合并。
- 合并只用于真正共享的分组。本例首列标题竖跨2行、服务笔数标题横跨4列；月份与合计仍是独立列。
- 用表头、合计浅底、必要水平线区分区域；先清所有边框，再逐格设需要的边。只有改笔颜色不代表既有边框已应用。
- 本机原生显示中，只设合计行top而上一行bottom=none，接缝没有显出期望强调线；将相邻两侧设为同宽同色后通过。边框两侧应一致并看原生PNG，不能凭XML里有top就宣布视觉正确。

## 固定样例入口

```text
python scripts/bili_layout_table_build.py --officecli officecli --out ./new/table.pptx
python scripts/bili_layout_table_verify.py ./new/table.pptx --report ./new/structure.json
powershell -NoProfile -File scripts/bili_layout_table_export.ps1 -Deck ./new/table.pptx -Evidence ./new/native
```

PPT、报告和导图目录都用新路径，入口拒覆盖；生成后显式close/flush并核磁盘中的真实table，再交外部PowerPoint。导图器只关闭自己的文稿，不Quit共享应用。

该样例为原创演示数据：线上100/120/130、现场70/80/90、伙伴30/40/40；行合计350/240/110，月合计200/240/260，季度700。替换业务数据时同步重算行列/总合计、标题、单位/口径及独立校验预期，不将演示数字当事实。PowerPoint table的数值是可编辑单元格文字；合计由生成器计算和独立验证，不承诺编辑一格后自动按Excel公式重算。

## 已验证范围

真实6行5列table、30个原始tc、26个视觉单元格；gridSpan4、rowSpan2及hMerge/vMerge续格完整。22pt东亚字体、右对齐、列宽/行高、必要边框、无竖边、单位来源与700合计均独立通过。PowerPoint实际HasTable、Rows/Columns及合并原点宽高比4/2通过；最终原生PNG已看，分组清楚、文字完整、合计强调线可见。

这不是原先1×5半透明排版表的复用证明；合并、数据与合计、边框范围和原生几何另行测试。更多行列、复杂交叉合并和其他表格样式需要按实际数据重新核验。
