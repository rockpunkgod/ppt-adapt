# 原生Section Zoom：进入一节、自动入口与节末返回

backend：`hybrid`。基础页、稳定对象名、Morph、Zoom基架由OfficeCLI生成；本机PowerPoint建立真实节并验证；最终SectionZoom XML通过OfficeCLI raw-set写入。已完成新作品、原生静态、实际点击与桌面录屏闭环。

来源：[One Skill PPT，Slide Zoom Template完整教程](https://www.youtube.com/watch?v=NFfgrwNsU9I) 11:48–13:46：Insert Section Zoom、Return to Zoom、关闭Zoom Background。节首是大图标，autoAdvance0，再Morph到正文，节末返回。这里处理的是Section Zoom，不能用Slide Zoom或普通链接替代。

## 配方与边界

1. OfficeCLI先做Hub和每节全部页面。示例五页：Hub1、A入口2/内容3、B入口4/内容5；节首advTm0/advClick0，内容Morph，图标跨页用同一`!!`稳定名。
2. PowerPoint `SectionProperties.AddBeforeSlide`建立真实节，读取`SectionID`，保存每节范围。section GUID不是首张slideId。
3. OfficeCLI `--type zoom`在1.0.160只生成Slide Zoom。可借它有效的graphicFrame、Choice/Fallback和媒体关系作为基架，但必须按官方规范把Choice的namespace/graphicData URI改为sectionzoom，节点改`sectionZm/sectionZmObj`，`sectionId`指真实节，`zmPr`也属于sectionzoom namespace。清除原SlideZoom的sldId/cId属性。
4. 同一zmPr设置`returnToParent=1`、`showBg=0`、`transitionDur=700`。图标海报从原生节首PNG生成，用hidden图片carrier建立有效r:embed，并同步Choice/Fallback。默认细灰框用原生spPr/ln/noFill去除；透明背景不是把整张海报调透明。
5. 独立读ZIP检查有真实sectionZmObj、GUID节范围、零sldZmObj；schema/issues通过。原生点击必须进入节内容，并从节末返回Hub。旧软件Fallback只是静态海报/单页链接，不能承诺节范围返回兼容。

规范依据：[Microsoft sectionzoom schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/55c6a539-eea9-478f-a6ff-6ad2f1dde7e7)、[ZoomObjectProperties schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/84585d01-dfd5-403e-9235-27853aff4d46)。本机旧PIA15没有Shapes.AddZoom方法，不猜`AddSectionZoom`接口。

## 包内执行入口

所有输出显式指定且拒绝已有目标文件/目录/报告；Python仅标准库调用OfficeCLI。下面的NEW_DIR等均须换成新路径。COM脚本只关闭自己的presentation，不Quit共享PowerPoint；共享桌面测试须与其他任务串行。

```powershell
python scripts/section_zoom_build.py --officecli OFFICECLI_PATH --outdir NEW_DIR
powershell -NoProfile -File scripts/section_zoom_prepare.ps1 -InputDeck NEW_DIR/section-base.pptx -OutputDeck NEW_DIR/section-with-ranges.pptx -EvidenceDirectory NEW_DIR/ranges
python scripts/section_zoom_convert.py --officecli OFFICECLI_PATH --input NEW_DIR/section-with-ranges.pptx --output NEW_DIR/section-final.pptx --evidence NEW_DIR/ranges
powershell -NoProfile -File scripts/section_zoom_record.ps1 -Deck NEW_DIR/section-final.pptx -Report NEW_DIR/native-record.json -OutputDirectory NEW_DIR/native-recording -Ffmpeg FFMPEG_PATH
python scripts/section_zoom_verify.py --deck NEW_DIR/section-final.pptx --click-report NEW_DIR/native-record.json --output NEW_DIR/independent-checks.json
```

`section_zoom_test.ps1`是只导静态图+真实点击的替代验证入口，参数Deck、Report、OutputDirectory同上；不需要录屏时使用。`section_zoom_record.ps1`含同样四项点击，并从DWM物理窗口边界裁切gdigrab desktop录屏。不要用title窗口捕获：本机该方式只录到Hub缓存画面，不能证明动态。

## 包内全新前向验证

2026-11-06，`validation/forward-section-zoom/`完全由上述包内脚本新生成，没有复制旧成品。5页schema通过/issues0，原生5张PNG逐张查看；独立结构+交互审计8项通过（independent-checks.json）。默认灰外框已去除并原生复看Hub。

`native-record.json`四项实际物理点击均通过：

| 操作 | 原生实际结果 |
|---|---|
|Hub点击A|自动经过节首2到内容3|
|内容3 Next|节末原生返回Hub1|
|Hub点击B|自动经过节首4到内容5|
|内容5 Next|节末原生返回Hub1|

点击不是GotoSlide代替；GotoSlide只在每轮初始复位Hub。测试关闭Presenter View，定位screenClass，用ClientToScreen+SetCursorPos/mouse_event。点击后等3.2秒覆盖Zoom和Morph；最初1.1秒等待会让Next落在过渡内被忽略，不能据此误判返回失效。HWND在本机COM可能读null，脚本保留EnumWindows定位，不把PIA属性存在当运行成功。

`native-recording/section-native.mp4`为13秒真实桌面裁切录屏，八个中间样本已看：0.2s Hub→1s连续放大→2/3s A正文→5.5s Hub→7.5s B正文→9s回缩中间态→11s Hub。不是只有首尾页截图。Hub薄荷底、入口米色；关闭Zoom背景后Hub圆标四周仍薄荷，原生播放正文期间也沿用Hub底，确认showBg0生效。

本配方通过的是两节进入/透明背景/0秒入口/Morph/节末返回。示例文字、配色、圆标均原创简化，不声称完全复刻作者全部31分钟作品。旧SlideZoom实验只验证单页返回，不能替代此处范围证据。
