# 悬停预览与多页 Section Zoom

backend：**hybrid**。复用[完整悬停菜单](hover-zoom-menu.md)的OfficeCLI页面、真实双孔遮罩、MouseOver状态及原生海报，再用[Section Zoom](section-zoom.md)的真实节GUID机制组织多页正文。没有改写这两项原配方的语义或验收标准。

## 页面与对象

7页：home1、hoverA2、hoverB3、Research4–5、Delivery6–7。先规划三种菜单状态和两节全部正文，再建立每节入口。悬停缩小图标并露出预览；点击预览进入一整节；两页讲完后原生返回该hover状态，移出卡片后回home。两条分支均可再次进入。

三个菜单页中的六个Zoom都改成真实`sectionZmObj`，分别绑定4–5/6–7的节GUID，不能只修改旧Slide Zoom的首张目标。两节都保留正文手动推进，Section Zoom时长500ms，悬停Morph350ms；正文使用原生UncoverRight（2051，XML为`pull dir=r`），不把正文的翻页误当成另一条Zoom。

转换必须原位替换Zoom的Choice，保持原始层级：Zoom和label在双孔遮罩后，标题在遮罩前。若转换时把新Zoom追加到遮罩上方，悬停过渡会泄露原本应隐藏的预览。为新章节首张重新导出原生PNG，同时关联Choice和Fallback，不能沿用改稿前的旧页脚。缓存图片保持hidden，不参与命中。

## 制作

先执行原有`hover_zoom_build.py`、`hover_zoom_polish.ps1`、`hover_zoom_posters.py`三步，得到五页`hover-final.pptx`。然后：

```powershell
powershell -NoProfile -File scripts/hover_section_prepare.ps1 -InputDeck NEW/hover-final.pptx -OutputDeck NEW/ranges.pptx -Sections NEW/sections.json -Posters NEW/posters
python scripts/hover_section_convert.py --officecli OFFICECLI --input NEW/ranges.pptx --output NEW/hover-section.pptx --sections NEW/sections.json --posters NEW/posters
```

适配器只接受自身五页新底稿，输出三菜单/两节各两页的示例；其他页数和内容应按实际演示规划调整。所有输出使用新路径。需要OfficeCLI、Python标准库、Windows PowerShell和桌面PowerPoint，不下载素材或调用另一个演示引擎。

## 组合实播

`examples/hover-section-zoom/`保留7页可编辑PPTX和约71秒连续原生录像。最终350ms版本独立完成两分支各两轮：25个稳定状态，每状态六次页码读回；30段连续相对鼠标MOVE、2910个轨迹样本、12次标记点击，无额外鼠标点击或按键。没有用GotoSlide、Next或链接API代替真实输入。

45个分布画面已复看，覆盖悬停变化、预览孔内保持、真实放大、第二页Uncover、节末回到对应预览、移出后回home及重复进入。六个Choice/Fallback海报都与最终原生目标导图逐字节相同，所有Zoom仍在真孔遮罩后。OfficeCLI验证零错。图像分析使用10fps重采样，不宣称原生录屏恒定帧率。

验证使用所有幻灯片（`RangeType=1`）；指定范围对原生节末返回的影响见[双首页](interaction-patterns.md)。测试的节末点击落点在返回后的卡片范围内，因此能停留在hover状态，再移动到背景验证home返回。不承诺指针已在背景时仍停留预览、任意极速横扫、Kiosk或其他阅读器的同等行为。首轮被其他窗口打断的记录作废，没有用于通过结论。

课程操作依据：[多页章节与悬停组合](https://www.youtube.com/watch?v=9Ewb0Rd3P3E&t=9648s)，2:40:48–2:43:28；已有本地原始转写及七个操作画面复核，网页任务空间已结束。9706秒画面的Zoom时长仍为默认1秒，最终500ms依据后续明确操作，不把中间态当最终值。作品使用原创两卡内容，不附作者模板或照片。
