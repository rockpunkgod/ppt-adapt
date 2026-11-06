# 能力范围与输入

本技能包用于制作、优化和验证可编辑PPT。按具体对象关系、动效和交付环境选择配方，适用范围以各配方为准。

## 输入

- 内容：主题、受众、主要结论、页数与信息关系。
- 素材：调用方可使用的图片、字体、音视频、数据及必要的PPTX/GLB模板。
- 环境：OfficeCLI、所需Python依赖，以及支持目标原生能力的桌面PowerPoint。
- 输出：目标路径、是否保留脚本、是否需要播放或交互验证。

## 使用边界

原生对象、原生动画和外部生成的图像必须按实际实现区分；保留原文、图表数据与必要的源素材。静态画面通过不代表动画通过，关系文件存在不代表媒体能播放，字段回读不代表实际点击成功。

只有列明并完成适用验证的配方作为可用入口。未实播的新组合不作为已验证能力；3D四圆点鼠标导航的独立通过范围见model3d.md，不推广为任意新模板。原生透明GIF裁形纹理的独立通过范围见[native-gif-texture.md](native-gif-texture.md)。新素材、新参数和不同软件环境需要重新验证。

必要素材与依赖不随示范效果自动取得授权；缺少输入时明确说明，不能用静态图片或近似对象冒充要求的原生效果。

笔尖与Ink Replay同步的独立通过范围见[ink-pen-sync.md](ink-pen-sync.md)：原创三段笔迹、透明圆心、同长路径与Fly前后串联，含真实返回后重新播放；不推广为任意自由笔迹自动同步。

单文本框逐字Wave循环的已实播范围见[text-wave.md](text-wave.md)：0.5秒/2秒、letter10%、两竖向幅度，保留原生旋转；不扩大UntilNextClick或任意文本支持。

Spin UntilNextClick与后续Fade的已实播范围见[until-next-click.md](until-next-click.md)：当前相位保持、两次Left与重新启动，另有UntilEnd对照。不推广到任意动画或指定对象触发。

同页GIF指定Appear与视频Fly/Play/Transparency的独立组合见[media-shape-triggers.md](media-shape-triggers.md)：两入口和空白对照已实播，保原媒体与70%透明；不推广为任意反复触发、跨页重入或精确GIF首帧。
