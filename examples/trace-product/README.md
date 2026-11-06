# 循迹 TRACE：团队知识工作台

原创概念产品的9页现场介绍。面向团队负责人和产品团队，讲清“从找到资料，到带着证据推进工作”。产品、界面与情景数据均为设计演示，不代表在售软件或实测性能。

艺术方向：编辑式科技产品演示。墨绿近黑、浅灰和酸橙色构成共同语言；Noto Serif SC用于展示标题，Noto Sans SC用于正文与界面。主体采用可编辑产品界面、文字构图、原生数据图和关系图，不依赖无关科技背景。

|页|核心信息与主体画面|节奏|
|---|---|---|
|1|循迹是什么：大标题与具体产品界面|安静开场|
|2|资料散落，找到文件还未接上工作：文字对比与分区|短淡化|
|3|先找回证据：宽幅产品工作台|产品主体|
|4|再接回工作：同一工作台缩移，行动区域出现|一次有目的的Morph|
|5|从“查到”到“接着做”：章节文字与工作流|停顿转折|
|6|把时间留给判断：原生比较图与演示差值|两次按系列揭示|
|7|控制权留在团队：权限、引用、人工确认关系|静态解释|
|8|从一个问题开始：小范围试用与阶段安排|强调色章节|
|9|将下一次搜索变成工作的起点：品牌收束与行动|保持结尾|

复用现有字体层级/字距、网格对齐、原生形状与分组、连续状态Morph、原生图表按系列构建和COM时长精修。制作使用OfficeCLI；预览与播放使用本机PowerPoint。未新增配方引擎、审美评分或兼容性数据库。

作品：`trace-product.pptx`。文字、界面、关系图和图表保留原生对象；图表用演示数值，不承诺嵌入Excel工作簿编辑。换机器需安装Noto Serif SC与Noto Sans SC，或替换字体后重新检查折行。

已用本机PowerPoint逐页预览并实际放映，确认界面连续缩移、图表两次揭示和结尾保持。[整套预览](preview.jpg) · [最终原生播放记录](playback.mp4)。

制作源为[build.py](build.py)，仅编排现有OfficeCLI原语。重新制作时在仓库根目录运行，输出用新路径：

```powershell
python examples/trace-product/build.py --officecli PATH-TO-OFFICECLI --out NEW-BASE.pptx
powershell -NoProfile -File officecli-ppt-design/scripts/polish_timing.ps1 -InputPath NEW-BASE.pptx -OutputPath NEW-FINAL.pptx -PlanPath NEW-BASE.timing.json
```
