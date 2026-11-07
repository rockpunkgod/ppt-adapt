# 一键重建与预览（Agent / 维护者）

普通用户只描述需求，Agent自行读取本页并处理运行环境。此入口从本地源文件重建首页当前九页PPT Adapt介绍，不复制现成演示，也不把其艺术方向强加给其他需求。

仓库用户可双击根目录`rebuild-demo.cmd`。默认在当前目录的`output/ppt-adapt-日期时间/`生成：

- `presentation.pptx`：重新生成的九页可编辑演示，包含原生纸张结构、文档空间与连续Morph。
- `preview.jpg`：PowerPoint原生渲染的整套预览。
- `preview.mp4`：PowerPoint原生视频导出，默认每页停留五秒，保留转场。

安装单个Skill后，同样可直接执行技能内部入口，路径相对本技能根目录：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/rebuild_demo.ps1 -OutputDirectory "D:/Presentations/demo"
```

运行条件：Windows桌面版PowerPoint、Python 3.10+、可工作的OfficeCLI。字体使用Noto Sans SC、Noto Sans SC Black、Noto Serif SC及Bahnschrift SemiCondensed；本机缺失时应先处理字体或检查替换后的折行。脚本不联网、不下载、不覆盖已有输出目录、不关闭用户演示文稿。

入口自动寻找Python以及仓库`tools/officecli.exe`、`PPT_ADAPT_OFFICECLI`和PATH内的OfficeCLI，并实际探测创建文件能力。Agent也可传入`-OfficeCli`和`-PythonExe`定位已有工具。`-SkipVideo`只输出PPTX和静态预览；`-KeepWork`保留本次构建素材以便调试，默认删除本次内部中间产物。失败时保留工作目录供Agent排查，不宣称成功。

内部顺序：用既有`build_components.Deck`和示例源创建方向页及产品场景 → 原生导图 → 重建九页 → COM精修 → 原生预览。主体由可编辑原生形状和路径构成，作品缩略图来自本次实际原生导图；不调用图像生成或下载服务。纸张展开、文档聚焦和页面展开使用已有Morph连续对象机制；二维空间表达不称为真实三维运镜。

需要复核当前作品的实播时，内部`demo_source/record.ps1`接受`-Deck`、新建的`-Evidence`目录及本机`-Ffmpeg`路径，原生放映并记录8次真实前进键。运行期间PowerPoint需保持前台；发现焦点丢失或页序不符即停止，不将录屏或导出成功单独当作视觉通过。它只适用于当前九页作品，不替代特殊配方原有交互验证。

任意新作品的预览可复用：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/render_preview.ps1 -InputDeck "D:/Presentations/result.pptx" -OutputDirectory "D:/Presentations/preview" -Video
```

预览目录可以是新目录或成品所在目录，但不能已有`preview.jpg`、`.preview-pages`或本次要导出的`preview.mp4`，避免覆盖文件。视频导出属于原生渲染预览，并非人工操作的实播放映录像；不把导出成功写成动效实测通过。新组合继续执行原配方要求的必要验证。
