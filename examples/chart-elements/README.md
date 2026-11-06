# 逐柱讲解：两种阅读顺序

`chart-elements.pptx`包含两页相同原生柱形图，用不同顺序组织讲解：第一页先展开Alpha四个季度，再展开Beta；第二页逐季度比较Alpha和Beta。每根柱子以0.5秒自底部擦除出现，保留真实chart、文字和原创示范数据。

`playback.mp4`为本机PowerPoint的37秒连续播放。第一页推进八次，再前进一页，第二页推进八次；17次实际方向键、无额外输入。两页18个稳定状态的逐柱显隐，以及16次自底向上的运动分别核对。`preview.jpg`展示关键运动的早、中、晚状态。

使用现有`chart_build_build.py --elements`重建，详见[制作方法](../../officecli-ppt-design/references/chart-build.md)。复用本机OfficeCLI即可，不下载素材。图表采用字面量数据，可由OfficeCLI改值；本示例不包含PowerPoint“编辑数据”的嵌入工作簿。改变系列、类别或另存后，需重新核动画索引和实际推进次数。
