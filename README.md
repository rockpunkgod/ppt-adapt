# PPT 制作技能包

制作、优化和验证可编辑 PowerPoint 演示文稿。包含版式、字体、配色、图表、母版、Morph、路径动画、交互导航及 3D 场景配方。

入口：[officecli-ppt-design/SKILL.md](officecli-ppt-design/SKILL.md)；按用途查找：[配方索引](officecli-ppt-design/references/recipe-index.md)。

## 使用

让支持 Skills 的 Agent 加载 `officecli-ppt-design`，然后提供演示主题、内容、页数、受众及素材。例如：

> 使用 officecli-ppt-design，制作一份 8 页的产品介绍 PPT，采用简洁版式、清晰图表和适量 Morph，保留可编辑文字，并用本机 PowerPoint 验证。

将整个 `officecli-ppt-design` 文件夹添加到技能库；保持 `scripts`、`references`、`assets` 和 `agents` 的相对位置。

## 实现方式

| 类型 | 用途 |
| --- | --- |
| `officecli-native` | 页面、排版、形状、图表、基础动画及 Morph |
| `powerpoint-com` | 原生动画的时长、延迟、加减速和行为精修 |
| `hybrid` | OfficeCLI 搭建基础，再用 PowerPoint COM 或必要的原生界面完成效果 |
| `external-render` | 核心图像或视频效果由外部计算或渲染生成，再嵌入 PPT |

## 环境

- OfficeCLI；本地验证使用 Windows 桌面版 Microsoft PowerPoint。
- Python 3 与 Windows PowerShell；按配方准备 NumPy、Pillow、SciPy 等依赖。
- 录屏和媒体分析需要 FFmpeg/FFprobe；字形几何配方另需 `scripts/requirements-geometry.txt`。
- 图片、字体和模板由调用方提供并确认可用。不同配方的输入限制与验证范围见对应说明。

技能包保留可复用脚本、必要素材及使用说明，不包含试验稿、录屏、研究缓存或临时报告。未通过的效果不作为可用配方。
