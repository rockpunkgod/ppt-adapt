# PowerPoint COM / VBA 动效精修

用户指定的工作链：AI理解源教程 → OfficeCLI基础制作 → PowerPoint COM/VBA精修 → 本机原生播放与录屏对照。COM 是正式制作层，不仅是验证器。

## 职责边界

- **OfficeCLI：** 页面与排版、原生文本/图片/形状、组合、基础动画和Morph、内容批量替换。
- **COM/VBA：** 精细Timing、加减速、重复/反转、指定对象触发与特殊动画等实际暴露的原生能力。
- **原生播放：** 验证中间轨迹、节奏、连续性和实际点击路径；截图只能验终态。

每次精修保留基础文件、新输出文件、可重放脚本与参数计划。不得只在应用里改好一次却没有复现路径。若原生COM也不暴露某功能，再考虑最小OOXML补丁或保留待实现状态，不改用近似名字掩盖。

## 精修现有动画，避免破坏整个时间轴

以页码、对象名称、该对象的第几个Effect定位，修改 `Effect.Timing`；不要按全局序号猜某个Effect属于谁。对复杂交互组还需区分 `MainSequence` 与 `InteractiveSequences`。

原生COM时间使用**秒**，OfficeCLI动画参数使用**毫秒**。`duration=1500` 的CLI值对应 `Timing.Duration=1.5`。加减速为0–1的比例；自动反转让完整周期为前向时长的两倍。重复三次不能只导出一个周期的视频。

PowerShell显式强转必须写`[single]0.5`，不要写`[single].5`；后者会按类型成员解析为null，传给COM可能成为0或最小1ms。Pin/Wheel曾仅状态通过而Fade实际1ms，已修复并新增500/400ms原生XML与影片审计。不要把这种语法错误误判为API不支持，也不要仅看最终开/关状态验证时间。含中文的Windows PowerShell脚本需使用可正确识别的UTF-8 BOM或明确的UTF-8读取，防止默认代码页把文字改成乱码。

`scripts/polish_timing.ps1` 是精修现有MainSequence动画的通用入口，默认拒绝覆盖目标。计划文件示例：

```json
[
  {"slide":5,"shape":"!!trail-0","effect":1,"duration":3,"delay":0,"autoReverse":true,"repeatCount":3},
  {"slide":5,"shape":"!!trail-1","effect":1,"duration":3,"delay":0.07,"autoReverse":true,"repeatCount":3}
]
```

例子的对象名要替换为成品真实名称；它不负责创建不存在的效果，也不把普通onClick变成指定形状触发。

## 先后与事件

- `withPrevious`、`afterPrevious`、普通页面点击分别表达同拍、串行、用户推进，不可互换。
- “点击Pin显示标签，再次点击隐藏”必须是指定形状的交互序列，并测试点击其他区域不会触发。
- 缓动和Bounce End是不同参数；原生API不可写时不能用独立bounce入口假装Fly末段反弹。
- 同一路径上递增延迟的对象，不一定能通过逐对象repeat让整组无缝循环；必须看下一周期是否保持相位关系。

## 验证记录应包含什么

1. 源视频时间段与目标轨迹/节奏。
2. 基础文件与精修后文件的差异；读回真实COM属性或OOXML确认写入。
3. 全部状态的PowerPoint原生PNG。
4. PowerPoint放映/导出的视频：转场前、中、后，以及足够长的循环段。
5. 交互测试起点、点击对象、实际目的页和返回页。
6. 对照后发现的问题、修正、复测结果。

多个agent只并发生成与分析；同一台PowerPoint的放映/导视频排队。仅关闭本次打开的presentation，不退出共享应用。拍到桌面或其他文稿的截图不可用作验证证据。

本机曾出现PowerPoint自己的“可能导致不稳定，请保存并重启”对话框，同时有Boolean命令禁用和真实输入失效。此时先停止并发原生操作，保存失败证据、关闭各自稿件并检查所有文稿/放映，不能继续用重复失败判定接口不支持。只有统一维护已确认应用完全空闲、没有用户文稿时才考虑正常重启；生产配方不自动Quit或强杀共享PowerPoint。本轮维护同时换了进程并补齐新子进程的CommonProgramFiles路径，部分同稿复测恢复，但不能据此断定单一原因或保证后续输入稳定。
