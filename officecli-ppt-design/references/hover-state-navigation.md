# 鼠标悬停切换介绍状态

backend：`hybrid`。OfficeCLI创建完整内容与五张状态页，PowerPoint COM建立原生MouseOver内部动作。普通点击导航、指定对象动画触发、鼠标悬停是三种不同事件。

来源：[BV1XS421R79e](https://www.bilibili.com/video/BV1XS421R79e/)43–90秒，61秒菜单整帧明确“鼠标悬停→超链接到幻灯片”。教程用人物照片和四份介绍；本配方用四张原创文字卡与原创介绍，保留真实事件和状态机制，不捆作者照片。全长可用字幕与六张分布整帧交叉核对，不声称逐秒连续观看。

## 机制

1. 先创建overview与四张选中状态页，最终页ID固定后再绑定动作；不沿用教程先建空白页、删掉再复制的易失链接步骤。
2. 每张未选中的卡：`ActionSettings.Item(2)`设`Action=7`，Hyperlink.Address为空，SubAddress绑定实际SlideID与SlideIndex。点击动作Item(1)明确为None；保存后必须是真正`a:hlinkHover`及内部slide关系。
3. 当前选中卡不设置到自己的悬停动作，避免循环跳回自身。四张状态页都保留其他卡入口。空白处无动作，移出后保持当前介绍；**不含移出自动回首页、点击进入Zoom或同页动画触发**。
4. 介绍正文与卡片标题分开排版，保留完整可编辑文字。所有页禁普通空白点击推进/定时推进；本配方没有跨页动画，状态即时切换。

## 可复现入口

Python基础构建和OOXML检查仅标准库；基础构建依赖同包`build_components.py`。Windows本机PowerPoint用于动作与文本边界检查，真实放映录屏另需FFmpeg。输出拒覆盖。

```powershell
python scripts/hover_cards_build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/hover_cards_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Evidence NEW/native
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/hover_cards_test.ps1 -Deck NEW/final.pptx -Directory NEW/live -Ffmpeg FFMPEG
python scripts/hover_cards_verify.py NEW/final.pptx --live-report NEW/live/report.json --report NEW/checks.json
```

放映测试仅在开始前调用一次GotoSlide回到基线。之后只有真实鼠标移动，没有任何按钮注入或COM翻页：依次卡1/2/3/4/1/4/2，每次移入后核页码和四个卡片背景像素，再移至空白处做保持状态负对照。仅定位自己可见screenClass，前台、WindowFromPoint根窗口和实际cursor坐标均守卫；DWM物理区域录制，不退出共享PowerPoint。

## 已通过与修正

`validation/forward-hover-cards/`为包内全新前向：CLI五页→COM16条Hover链接→五页原生PNG及四份正文边界→15真实状态（七次移入、七次空白、初始态）→12秒实际录屏→53条独立结构/状态检查→OfficeCLI validate全部通过。全部介绍均完整位于卡内。

早期一体式20pt长文在卡外溢出，已改独立17pt正文并以原生BoundTop/BoundHeight拒绝溢出，未删原介绍。输入测试也曾失败：只SetCursorPos（含零delta MOVE）不能稳定触发本机会话的Hover；最终先定位，再发送相对+1/-1鼠标MOVE、等待50ms处理事件，随后核实际坐标逐点回原处。没有注入点击，也没有用GotoSlide冒充Hover成功。此处是实际输入验证工具的修正，不是PPT播放必须依赖脚本；手动播放仍使用原生悬停动作。

原空自动化进程曾正式报不稳定，维护后初次旧稿通过，但后续仅坐标输入仍失败，故不能把输入恢复全部归因于重启或某一环境项。保留失败记录，不从命令读回成功推导交互成功。

边界：当前四卡/五状态与所列七次切换；不声明任意连扫、全部显示器配置、触控/笔输入、网页版或非PowerPoint阅读器行为。修改卡片位置、文本或目标页后重跑原生边界与真实悬停检查。
