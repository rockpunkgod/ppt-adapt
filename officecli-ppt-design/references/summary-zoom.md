# 真 Summary Zoom：原生摘要容器与自动网格

后端：hybrid。OfficeCLI 制作原创基础页；PowerPoint 可见 UI 创建/编辑真实 Summary Zoom；COM 仅节设置、几何尺寸、导图与保存；严格物理点击验证。

源：[PPT完整大师课](https://youtu.be/9Ewb0Rd3P3E) 1:30:15–1:32:01，原转写677–690；源5425s插入对话框、5461s窄容器、5500s四列。这里只验三节各两页的最小工程，不把整个章节的所有交互变体标通过。

## 执行入口

按顺序运行本参考配套的 `summary_zoom_build.py`、`summary_zoom_author.ps1`、`summary_zoom_finish.py`、`summary_zoom_record.ps1`、`summary_zoom_verify.py`。另须同目录保留共享 `summary_zoom_ready.ps1`：作者与录屏均检查原生导图中2/3种原创章节色实际出现，最长4秒有界等待，不接收灰色占位缩略图。每个步骤以同一个新的输出目录保存其独立证据。所有现有目标均拒覆盖，失败时换新目录保留负例。

1. build：`--officecli <officecli.exe> --out <new-directory>`，生成六页原创作品。三章为 Field Notes / Signal Lab / Next Route，各两页；调用方可在保持结构后替换原创内容。
2. author：`-Deck <out>/summary-base.pptx -Out <out>`。必须在独占 PowerPoint UI 槽内执行；原生 Insert → Summary Zoom 选择第1/3页，并保留未用节；随后宽/窄改变由PowerPoint自动重排；原生 Edit Summary 加入第三节。该步骤绝不读取预制 Summary 模板或复制 seed XML。
3. finish：`--officecli <exe> --out <out>`，仅给原生摘要页加独立标题/页脚，解析实际 section/slide IDs 和原生子项坐标生成点击计划。
4. record：`-Deck <out>/summary-final.pptx -Plan <out>/click-plan.json -Evidence <out>/live -Ffmpeg <exe>`，独占UI，保存实际桌面录屏、原生PNG与9次物理输入日志。
5. verify：指定 `--initial <out>/summary-sectioned.pptx --edited <out>/summary-edited.pptx --deck <out>/summary-final.pptx --resize <out>/resize-probe --plan <out>/click-plan.json --live <out>/live --ffmpeg <exe> --out <new-audit-dir>`。需 Pillow/NumPy；查看输出四张接触图，逐项确认全片缩放及内容。

目前只验证本机中文 PowerPoint 的实际可见控件。作者流程使用观测到的 InsertZoomMenu / UpdateInteractiveIndex ID、中文标签与对应 UIA Pattern，不猜 ExecuteMso 命令。控件缺失、对话框类型不匹配或前台守卫失败即停，不能跳过该步骤后声称创建成功。

## 对象与验收

真对象：`graphicData` URI `http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom`，内部 `summaryZm`、section GUID、原生 `gridLayout`。fallback 是兼容用链接图片组，不能用它代替主动对象。此对象本机COM Type=-2、GroupItems.Count=0；因此必须同时检查实际XML、上下文Edit Summary UI和原生行为。

验收2→3真实子项、保留原A/B IDs；宽度90%为两列、38%为一列，PowerPoint自行改变子项坐标并保持实际页面比例。九次真实点击需得到 [2,3,1,4,5,1,6,7,1]，包括三章的节末自动返回。

桌面放映先以自己 Normal 窗口真实 Caption UIA.SetFocus + WScript.AppActivate，再 Run Type1；只查自己 PID 的 screenClass，按实际 client/DWM 物理边界录制。不得向其他窗口输入、退出共享 PowerPoint 或关闭外部应用。若不能证明前景/鼠标命中归属，不得用COM Next/Goto代称物理点击。仅关闭自己的 show/文稿。

独立检查器比较每个桌面状态与原生导图的平色块和文字识别，另提取全片8fps接触图。参数/结构通过不代证视觉；完整录屏采样审阅不等于逐帧30fps观察。

## 保留边界

不等同于独立 Slide Zoom、手工 Section Zoom 拼组；未验任意章节页数/跨文档迁移/所有UI语言/所有版本。PowerPoint原生UI创建是必要作者环节，COM/OfficeCLI当前无受测从零Summary对象API。无需外部图片、下载作者模板或上传数据。
