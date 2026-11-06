# 笔尖跟随原生墨迹

backend: **hybrid**。OfficeCLI制作页面和可编辑铅笔，复用已有InkML与原生Replay种子，PowerPoint COM编排入场、Replay、自定义路径和离场。它不改变[5+2秒墨迹循环](ink-replay-pause.md)的机制或验收标准。

## 制作方法

1. 保留真正Ink对象及`drawProgress`行为。旧COM将Replay读成EffectType0，不能按0重建它。
2. 用无填充、无线条的圆覆盖铅笔，圆心放在笔尖，再整体组合。透明圆仍参与组的边界计算；路径移动的是组中心，因此中心须与笔尖重合。
3. 按笔迹绘制顺序建立自定义路径。本例三条等长直线，对应已有原创InkML的183个XYT采样；以页面宽高归一化路径位移。
4. Replay与笔尖路径同为1秒，路径WithPrevious，均无平滑开始/结束。不要给路径单独加减速，否则同起同止仍会在中途脱离笔迹。
5. 用AddEffect追加Fly入场，移到Replay之前：入场0.5秒、最大平滑结束；Replay改AfterPrevious，路径仍WithPrevious。最后追加0.5秒Fly Out/To Left，AfterPrevious、最大平滑开始。保留原Replay，不替换整条时间轴。
6. 先以5秒慢放看清位置，再恢复1秒完整组合。检查实际绘制过程，不能仅凭最终折线或移动的铅笔宣称同步。

示例提供两个内容页，另有两个隐藏的0秒复位过渡页。导航链接指向内容页前的过渡页，随后原生自动前进，使返回后的点击动画重新初始化。文稿须使用**排练计时/UseSlideTimings**；内容页保留单击推进。首次直接链接内容页能播放，但返回后会保留完成态；只关闭内容页的AdvanceOnClick又会阻止本例的单击动画。这是示例导航的组合修复，不是Replay本身所需的固定流程。

## 运行

```powershell
python scripts/ink_pen_build.py --officecli OFFICECLI --outdir NEW/base
powershell -NoProfile -File scripts/ink_pen_polish.ps1 -InputDeck NEW/base/before-native.pptx -Model NEW/base/model.json -OutputDeck NEW/ink-pen-sync.pptx -Native NEW/native
```

需要Python、OfficeCLI、Windows桌面PowerPoint与Windows PowerShell；示例字体为Noto Serif SC/Noto Sans SC。输出路径必须是新路径。临时静音只供已有Replay种子绑定，原生精修时删除，不承担最终播放控制。

## 已实播范围

`examples/ink-pen-sync/`保存可编辑PPTX和35秒原生播放。保存关闭后重开，六次真实鼠标点击完成慢放、进入快放、首次播放、返回、再次进入及再次播放；复看20个分布帧，并检查入场、离场和完整墨迹保持。

三段分别按30fps测量独立白色墨迹末端与粉色笔尖：慢放146个中间态，首次快放29个，返回快放28个；最大距离分别0.1428、0.1082、0.1143cm以内，绘制跨度分别20.30、19.77、19.00cm。2mm探针容差包含笔尖遮挡和抗锯齿，仅用于本原创示例。结构另确认真Ink、5秒/1秒drawProgress、无平滑路径和原生Fly参数；OfficeCLI验证零错。

任意曲线、长度不等的分段或停笔需要把InkML的T与笔尖进度重新对应；同一总时长不保证中途同步。没有验证任意自由手写、自动路径识别、网页播放器或其他PowerPoint版本。

同一基线曾在一个会话中静态显示；确认任务实例已无文稿/放映、正常退出，再以修正子进程环境和独立任务宿主重开后，同一基线恢复逐笔播放。多个条件同时改变，不能断言唯一原因；不得重启含有用户文稿的实例。遇到该情况应检查真实播放，而非修改已有Replay冒充修复。

制作依据：[完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)4:17:08–4:21:38。复读本地完整对应转写，并核看归档4:21:16操作帧中的To Left和0.50秒；本轮Ego页面断网、Watch字幕请求失败，未把失败重访计为新增观看覆盖。课程用自由笔迹，本例用原创三段笔迹演示同一组合机制。
