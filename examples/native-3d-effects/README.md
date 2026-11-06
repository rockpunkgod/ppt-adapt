# 原生3D动作

打开`native-3d-effects.pptx`：先单击空白，模型以原生3D效果进入并持续转动；点击模型，触发跳跃转身。文字和真正的3D模型保持可编辑。原创双面GLB已嵌入PPTX，也保留在`assets/card.glb`，无需下载。

页面让衬线大标题与右侧模型形成主次，深绿/酸橙贯穿TRACE视觉语言。右侧模型的完整跳跃空间优先于静态尺寸；缩小并下移避免最高点越界。正背面不同颜色及非对称标识帮助辨认真正的空间旋转。

`playback.mp4`是最终作品40秒本机PowerPoint播放；一次开始、三次模型点击，跳跃后仍持续转动。`preview.jpg`包含入场、跳跃高点及后续转动的分布画面。当前通过范围是这一单页组合，不覆盖跨页返回或快速连击。

在仓库根目录用新路径重建：

```powershell
python officecli-ppt-design/scripts/native_3d_effects_build.py --officecli OFFICECLI --model examples/native-3d-effects/assets/card.glb --out NEW/base.pptx
powershell -NoProfile -File officecli-ppt-design/scripts/native_3d_effects_polish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Officecli OFFICECLI
```

方法和接口边界见[原生3D动作指导](../../officecli-ppt-design/references/native-3d-effects.md)。模型制作复用已有`glb_card_build.py`；未改写既有Morph或3D导航配方。不同模型、比例或动效组合重新原生预览。
