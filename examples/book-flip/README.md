# 原生书页翻转

[可编辑PPTX](book-flip.pptx) · [原生播放](playback.mp4)

原创两页TRACE札记场景：曲线纸页从右侧绕书脊翻到左侧，露出下一页内容。书本、纸页、文字和流程图均为PowerPoint原生可编辑对象。

制作复用已有OfficeCLI形状/自定义曲线、组合、Morph及COM相机机制。可见右半页与水平镜像的辅助左半页先组成完整对开组，再隐藏辅助半页；组中心保持在书脊。相机PerspectiveRelaxedModerately=62、透视60°、RotationX从0到180°、Contrasting灯光18、Morph2秒。原生读回可将180°规范为−180°，表示同一终态。

隐藏辅助半页前后组的Left/Width保持，保存重开后仍能翻页。已用本机PowerPoint实播15秒，物理Right→Left→Right完成翻页、回退与再次翻页；复看14个分布帧，包括纸页侧缘、透视中间态、书脊稳定、内容逐渐露出和结束保持。OfficeCLI文件验证通过。制作中缩小了书页，为中间态的透视放大留出空间，避免压到页眉和提示文字。

背面为留白纸面。多张连续翻页、双面印刷文字、动态弯曲纸张及任意相机/尺寸不在本例通过范围。源的照片纸面以原创原生曲线替换；本例没有使用原作者素材，也不将这种平面组称为真正GLB模型。

在仓库根目录运行，输出使用新路径：

```powershell
python examples/book-flip/build.py --officecli OFFICECLI --out NEW/base.pptx
powershell -NoProfile -ExecutionPolicy Bypass -File examples/book-flip/finish.ps1 -InputDeck NEW/base.pptx -OutputDeck NEW/final.pptx -Native NEW/native
```

正文需Noto Sans SC与Noto Serif SC。改尺寸、内容、相机或支点后重新预览中间画面。
