# 指定对象触发 GIF 与视频叠加

backend: **hybrid**。OfficeCLI制作页面、蛋糕组合、标题和真实媒体，PowerPoint COM组合已有指定触发、入场、播放与透明强调。GIF和视频是准备素材；本机制不把粒子视频称为PowerPoint原生粒子模拟，也不替代[原生动画导GIF](native-gif-texture.md)。

## 页面与对象关系

同页建立两个独立入口：`cake`触发真正多帧GIF的Appear；`happy-birthday`触发视频Fly From Top、Play和Transparency。第二页保留相同场景，但省去透明强调，供比较视频透底效果。

先定背景，再放全画幅视频，其上放标题、蛋糕和GIF。视频位于底图之上、前景内容之下，70%透明时保留背景；GIF在视频上方，保持自身的亮度。让触发对象的有效点击区域露出，不能让GIF透明外框盖住蛋糕的全部可点击区域。本例点击蛋糕主体，标题点击文字本身。

GIF用真正Appear指定到Cake，不加媒体Play，也不改成普通下一次单击。视频先关闭自动播放，清理该视频自动附带的原生效果，再建立标题触发组：

1. Fly From Top，0.5秒，指定标题点击。
2. MediaPlay，与上一动画同时、零延迟。
3. Transparency 70%，与上一动画同时、零延迟，保持结束状态。COM `EffectParameters.Amount=0.7`落盘为`style.opacity=0.3`；不是形状填充透明度。

本例关闭页面空白点击推进，保留指定触发；空白不会启动尚未触发的媒体。不要清理业务稿的整条时间轴，示例脚本只接受自己的两个新基础页。

## 多触发组的接口处理

本机`MainSequence.AddTriggerEffect`拒绝创建；沿用已有机制，从`InteractiveSequences.Add(...).AddTriggerEffect(...)`创建。两个入口共存时，直接使用返回Effect的`Parent`曾误取GIF组，将视频Play和Transparency放错。应遍历实际交互序列，以视频对象ID和唯一Fly效果定位标题组，再直接向该序列追加Play和Transparency。不要凭“效果数量正确”宣称组合正确。

交互组中的WithPrevious效果，COM `Timing.TriggerType`也可能回读4；保存后的`nodeType=withEffect`、零延迟和实际一次点击同时执行更能说明归属。示例原生重开后再核对：Cake组只有GIF Appear；标题组为视频Fly/Play/Transparency，主序列为空。对照页标题组只有Fly/Play。

## 运行

```powershell
python scripts/media_trigger_assets.py --ffmpeg FFMPEG --outdir NEW/assets
python scripts/media_trigger_build.py --officecli OFFICECLI --assets NEW/assets --out NEW/base.pptx
powershell -NoProfile -File scripts/media_trigger_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/media-triggers.pptx -Native NEW/native
```

素材脚本以Pillow制作原创24帧透明烟花GIF、12秒H.264纸屑视频及首帧；视频直接通过FFmpeg编码，不留下逐帧图片。它们是可替换输入，不是核心制作引擎。需要Python/Pillow、FFmpeg、OfficeCLI、Windows桌面PowerPoint及Windows PowerShell。示例字体为Noto Sans SC、Arial。所有输出必须是新路径。

## 已实播范围

`examples/media-shape-triggers/`保留可编辑PPTX和两段原生播放拼接：主场景36秒、对照20秒。最终文件经保存关闭重开，七次真实点击分别核验空白、Cake、再空白、标题、播放后空白，以及不透明对照。窗口输入日志只记录到计划点击，30个分布帧已复看。早期出现计划外操作的对照片段未计通过。

点击Cake前烟花区域无亮线，显示后白色烟花的半径持续伸缩，重置间隔约2.4秒，覆盖多轮。标题前视频未显示、未被Cake或空白启动；标题点击后从顶部覆盖到全页，视频在GIF之外的区域持续运动，播放后的空白点击也未中断本轮播放。

透明页的背景平坦区域由录屏RGB约75/109/100变为52/76/70，保留约70%；不透明对照变为0/0/0，且同一视频继续运动。原视频、GIF嵌入字节保持，GIF为24帧、100ms/帧、无限循环、四角逐帧透明；OfficeCLI验证零错。不是用静态低透明图片、只有入场的海报或移动GIF代证视频播放。

已测的是列明两页、两触发入口和12秒素材。不承诺GIF精确从第0帧重启、任意跨页返回、反复点标题的状态机或所有播放器；换素材、层级和触发对象后须重新实际点击。蜡烛为静态装饰，本例不声明课程的火焰缩放组合已由此验证。

制作依据：[完整大师课](https://www.youtube.com/watch?v=9Ewb0Rd3P3E)4:21:53–4:25:41。Ego实际复核九个操作画面；15917秒仍为默认50%，15922秒才显示已输入70%，没有混淆中间态与最终值。Watch本地引擎取得197段英文原始自动字幕，完整对应内容已复读；网页空间已结束。使用原创素材，没有捆绑作者照片、GIF或股票视频。
