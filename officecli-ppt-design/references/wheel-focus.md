# 复合聚焦与 Wheel 环形相位

backend：`hybrid`。OfficeCLI制作可编辑基础圆、圆环与文本；PowerPoint COM精修原生动画和行为Timing。需要Windows本机PowerPoint，不捆绑第三方插件。

来源：[BV1VTDGYsEmd P22](https://www.bilibili.com/video/BV1VTDGYsEmd/?p=22)，1326–1737秒聚焦，1740–2277秒环形项目。完整研究为3秒字幕OCR样本、6分布帧和定点操作帧，不能称逐秒完整口述。原课试过多个透明度/环时长，配方明确采用透明50%、环7秒和等角六节点。

## 已通过配方

1. **聚焦**：三圆原位放置；1秒后三项同步GrowShrink59，中心Size120、两侧85，Duration1.5秒。两侧另加Transparency62，保留原生默认Amount0.5。必须同时将Effect.Timing及每个Behavior.Timing的RewindAtEnd设为false，才能在强调结束后保持透明。只设父层会在1.5秒后恢复，已由原生影片排除。透明是原生强调效果的Set，不声称连续渐变透明。
2. **环绘制**：中心Fade10 0.5秒；外环Wheel21 afterPrevious，7秒、FilterEffect.Subtype18（一辐条），实际产生`wheel(1)`。六节点按-90°+i×60°放置并Fade0.4秒，同withPrevious组Delay=i×7/6秒。Delay相对该并行组起点，不可每项都写1.167秒；那会五项同时出现。改角度、起点、轮辐数、时长后必须重新计算/观看，不能直接沿用。

## 使用

所有输出必须显式给定，已有输出/目录会被拒绝。以下`NEW`必须为新目录。脚本只依赖Python标准库、OfficeCLI与本机PowerPoint；原生验证关闭自己的文稿，不退出共享PowerPoint。

```powershell
python scripts/wheel_focus_build.py --officecli OFFICECLI --outdir NEW
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/wheel_focus_polish.ps1 -InputDeck NEW/wheel-focus-base.pptx -OutputDeck NEW/wheel-focus-final.pptx -EvidenceDirectory NEW/polish
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/wheel_focus_native.ps1 -Deck NEW/wheel-focus-final.pptx -OutputDirectory NEW/native
python scripts/wheel_focus_verify.py --deck NEW/wheel-focus-final.pptx --output NEW/checks.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/wheel_focus_record.ps1 -Deck NEW/wheel-focus-final.pptx -Report NEW/desktop.json -OutputDirectory NEW/desktop -Ffmpeg FFMPEG
```

## 验收与证据

exact包内脚本前向产物：`E:/ppt skills/validation/forward-wheel-focus-v2/wheel-focus-final.pptx`。独立XML审计8项全部通过，包括85/120/85、两项opacity0.5、子行为hold、五项1000ms同步、单辐条7000ms、六相位延迟。OfficeCLI基础validate/issues为0。

历史v1的`validation/forward-wheel-focus/native-v2/native.mp4`与12帧仅证明环相位及透明保持，不证明500/400ms Fade；该版实际Fade1ms。最终v2采用下述新桌面16秒放映与新增持久化Duration断言，旧影片不作最终时长证据。

这两页没有指定对象点击交互。另已补同一exact成品的真实放映窗口桌面录屏：`desktop/wheel-focus-live.mp4`16秒，`desktop/contact.jpg`12时刻已看，初态、缩放与透明保持、手动Next后环逐项绘出均通过。显式PointerType=1并回读，使用DWM物理窗口范围裁切desktop，不用title窗口缓存。该录屏同时验证实际放映而非仅CreateVideo导出。导图只展示编辑态终态，不能替代影片。建议复用时检查初态/运动中/结束保持/环相位四类时点。失败版和低对比版不入库；母版循环、人物光条组合不在此配方范围。

修正记录：早期`[single].5/.4`被PowerShell解析为null，核心/节点Fade实际1ms。现用`[single]0.5/0.4`，新增持久化500/400ms断言；v2新作品+8项审计+真实16秒放映已复看通过，使用v2证据。v1仅保留相位与透明保持观察，不再作Fade时长证明。
