# 时间、路径与动效精修

## T1 逐次点击揭示

来源：Top10 00:48–02:38。每个语义模块一个Fade入口，`onClick`，约500ms。PowerPoint放映中每次Next推进一个动画，仍留在同页。它不是点击指定Pin触发；后者进入独立交互序列。

`build_visual_animation.py`的animation-lab第1页为通过样例。动画序号、放映click index和实际目标分别验证；视频导出不自动替用户点击。

## T2 同拍起始，短延迟错开

来源：Top10 04:10–06:19。`withPrevious`配0/100/200/300ms递增delay形成瀑布；不是串行`afterPrevious`。

样例animation-lab第2页省略原作未验证的末段Bounce。看中间帧时，物体应处于不同位置；不能只检查终态全到位。COM读某一Effect的TriggerDelayTime可能遗漏外层容器延迟，以时间树和实播交叉确认。

## T3 原路径自动反转与循环

来源：Top10 06:19–07:43。一条路径的自动反转返回起点，完成一轮后重复。

路径坐标按画布归一化；`M 0 0 L 0.6 0 E`的位移是画布宽度的0.6，而不是形状宽度的0.6。自动反转使完整周期为前向时长的两倍。

本轮CLI的repeat/autoReverse存在“帮助声称支持、add忽略、set甚至重建动画”的实测问题。可用经过验证的局部OOXML写入，或用COM精修既有Effect；不得对复杂页宽泛匹配全部cTn。

样例animation-lab第3页的往返与重复已用原生视频确认。可用COM实现精修时优先保留基础作品与精修脚本，便于维护。

## T4 圆轨道上的光点

来源：SlideZoom25:13–28:59。源是柔边大圆+发光小圆组合，6秒匀速循环。

样例采用单发光圆简化美术；用四段三次贝塞尔构造圆，控制系数0.55228475。x/y分别除以画布宽高，初始**中心**落在轨道顶部。画布非正方形时不能对x/y写相同归一化半径。

animation-lab第4页原生视频实测采样点的径向误差在0.56像素内（5像素容差）。通过的是贴轨循环，不是源图完整复合光晕。

## T5 重叠文本，同一路径，递增延迟

来源：大师课05:54:42–06:01:30。7层相同文本起点完全重合，最前亮色、后层较暗，同路径和时长，delay0..420ms。

**基础：** `build_timeline.py`创建单程拖尾。**精修：** `polish_timeline.ps1`为7个既有运动Effect设置反转和3次重复，生成新成品。前向3秒，往返6秒，完整3轮至少18秒，另加最后层420ms及停留余量。

实测精修后PowerPoint可能重整delay所在的XML容器；不要仅断言某个固定层级的cond，既要读真实时间表也要看完整周期影片。若set操作把7条路径变成1条，必须从基础文件恢复再精修，不接受“命令Updated”。

## T6 加减速与特殊效果

用户指定COM/VBA为正式精修层。Duration、TriggerDelayTime、Accelerate、Decelerate、AutoReverse、RepeatCount作用不同；Timing使用秒，加减速是0–1比例。

Fly末段Bounce End与独立Bounce入口预设不同。指定对象触发和普通页面点击不同。只复用 [validation.md](validation.md) 明确通过的专用精修；未完成的原生Ink Replay与模型专属触发不由近似效果替代。

## T7 音频书签驱动整组循环

大师课5:03:05–5:05:24的同一时钟方法已独立还原，见[音频书签序列](audio-bookmark-sequence.md)。首条AddTriggerEffect放在InteractiveSequence，后续效果同组WithPrevious与各自delay；不是每条各建媒体触发。包内前向实播多轮通过，不能由此推导Ink Replay暂停循环、跨页音乐或透明视频能力。
