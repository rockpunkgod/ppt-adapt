# OfficeCLI 能力与实现边界

在本机 1.0.159 / 1.0.160 上研究与实测。`help` 用于发现语法；成品回读与本机 PowerPoint 放映用于确认真实行为。本文不会因为 help 声称支持就判定功能可用。

真实原生结构比相似API名称更可靠：多片连续图片应使用[父组blipFill与子grpFill](group-image-fill.md)，本机COM Group.Fill.UserPicture会各片重复整图；需要PowerPoint“编辑数据”时应补[真实工作簿绑定](chart-workbook-materialization.md)，不能由chart节点存在推断。两项都有原生画面与独立数据/像素对照。

## 选择执行层

1. 使用 `create/add/set/get/query/batch` 建立可编辑的文本、图片填充形状、组合与连续状态。
2. 使用 `add ... --from /slide[N]` 复制完整状态，避免重建时对象身份漂移。
3. 高层属性不能表达且有源证据时，用 `officecli raw-set` 做最小修改。随后 `close`、检查 OOXML、`validate`、本机 PowerPoint 验证。
4. PowerPoint COM/VBA作为正式精修层，补充精细Timing、缓动、重复、指定对象触发及本机暴露的特殊动画。基础页面仍通过OfficeCLI制作，区分两层贡献；用可复现脚本记录精修，不把COM实现冒称CLI原生支持。详见 [精修流程](powerpoint-polish.md)。

## 已确认语法

```text
officecli create deck.pptx
officecli set deck.pptx / --prop slideWidth=33.87cm --prop slideHeight=19.05cm
officecli add deck.pptx / --type slide --prop layout=blank --prop background=F3F6F8
officecli add deck.pptx /slide[1] --type shape --prop name=!!card-01 --prop geometry=roundRect --prop x=2cm --prop y=5cm --prop width=9cm --prop height=11cm --prop fill=174A64 --prop line=none
officecli add deck.pptx / --from /slide[1]
officecli set deck.pptx /slide[2]/shape[@name=!!card-01] --prop x=20cm
officecli set deck.pptx /slide[2] --prop transition=morph-700
officecli close deck.pptx
officecli validate deck.pptx
```

真实调用请将方括号路径作为单独参数；PowerShell/跨平台脚本推荐参数数组或 batch JSON，避免 shell 扩展。

## Morph 的状态管理

- 同一对象在相邻页保持相同、唯一的 `!!name`，目标页设置 Morph；画布独立，没有“对象自动存在于以后所有页”的机制。**复制旧页才可能带来旧对象残留。**
- 离场要连续运动时，在目标页保留同名对象并移出画布；普通消失可移除。不为从未出现在新页的对象虚构幽灵。
- 名称不能在同一可呈现分支内重复。`mc:Choice` 与 `mc:Fallback` 同一对象的 ID/name 重复不等于两份可见对象，检查时要识别分支。
- 只有一个对象运动也可以是有效 Morph，不为了满足“至少三物体”的人为规则加装饰。
- 所有对象的 Morph 同时发生；分先后使用中间状态页或真实对象动画。
- 对称图形的 rotation 不同仍可能绕一圈；源 Zoom 教程就出现此问题。两态不需要旋转时都显式归零。

## 本轮发现的陷阱

| 症状 | 已定位原因与处理 | 验证依据 |
|---|---|---|
| `--version/help` 成功，`create` 缺 System.Private.Xml | 本机一次安装的可执行文件异常；使用经实际 create 验证的副本。不要把 `--version` 当作就绪测试 | 研究目录 runtime probe；官方1.0.160实测 create 成功 |
| 链接把标题变蓝并加下划线 | 将视觉文字与链接命中区域分离，用透明热点承载链接 | 大师课 v1→v2 PowerPoint 原生截图 |
| `repeat/autoReverse` 声称更新但 XML 没有 | 回读结果和效应时间节点；必要时通过 raw-set 写入路径动画所属 cTn 的 autoRev/repeatCount，不改主时间轴节点 | 动画实验：COM读 AutoReverse、RepeatCount 及往返视频 |
| `add animation` 不接受 picture 父路径 | 对需要对象动画的图片优先用原生 shape 的 image fill；具体类型能力先看本机help并试验 | 三教程组的实际拒绝响应与替代作品 |
| `add zoom` 原生结构存在但缩略图空白 | 原始 blip 是1像素占位，不能指望 PowerPoint 自动生成；需嵌入目标页缩略图并修 Choice 与 Fallback 关系 | Zoom实验，修复后的放映结果见单独测试记录 |

**近似项不要混淆：** `bounce` 入口预设不等于 Fly 的 Bounce End；`trigger=onClick` 表示放映点击顺序，不等于点击指定形状触发；整页 Zoom 转场、Slide Zoom、Section Zoom 是三种东西；有 GLB 不等于带有模型内部骨骼动画。

## 各效果的数据结构

| 机制 | 优先实现 | 额外能力条件 |
|---|---|---|
| 卡片 / 轮播 / 时间轴 | native shape/group + clone + Morph | 图片填充、连续名称、裁切窗口 |
| 多笔刷擦除 | 裁切笔刷素材 + 独立 wipe | 素材要真透明；动画分笔排序 |
| 镂空字 / 异形孔 | 原生自定义路径 / 真透明遮罩 | 普通背景色文本不是真镂空；必要时预制矢量遮罩 |
| 模糊玻璃 | 预处理模糊图块 + 原生面板与文本 | 透明白矩形不能假装有背景模糊 |
| 3D内部动画 | 保留原 GLB 的动画轨道与 embedAnim | 资产必须可合法使用，时间轴与场景ID另验 |
| 交互导航 | 原生Zoom或内部链接 / 热点 | 目标关系、返回路径和播放器点击都要实测 |
| 触发器、Ink Replay、媒体播放 | 源中真实事件/墨迹节点 | 高层不支持时暂列待实现，不换名称冒充 |

“裁切笔刷素材”仍是候选路由；已独立通过的另一条是[原生笔刷背景填充与Wipe](local-blur-background.md)：清晰背景、虚化副本、真轮廓分笔擦入。它不依赖PNG裁切，也不把本次通过扩展到上表尚未实测的素材方案。

## 本机 PowerPoint 测试

在 PowerPoint 原生渲染全部状态页；开窗口式幻灯片放映，检查点击进入、返回、重复激活和任意位置点击是否误跳页。动态样例通过 `CreateVideo` 导出后，查看转场前、中、后帧；循环测试时影片长度须覆盖全部循环，不截一圈便宣称三圈通过。

所有并发生成可各用不同文件；PowerPoint COM 渲染/导视频排队，关闭本任务自己的 presentation，勿 `Quit` 共享 PowerPoint 或关闭用户文稿。截图避开桌面/其他文档。
