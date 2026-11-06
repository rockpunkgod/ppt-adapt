# 跨页Push时间轴接缝

backend：officecli-native。原理是把一条连续路线分配到相邻页面，前页底端与后页顶端使用相同x位置、线宽、颜色，再用向上Push把两页整体接起来。无需把它改成Morph，也不用复制多层离场图片。

来源：BV1VTDGYsEmd P20 1749–1812秒。完整可用3秒OCR已读，1760/1780/1795/1800/1804/1810及1767–1778逐秒整帧已看；1774是明确的向上Push中间态，1795–1804强调前页底部/后页顶部紧贴，没贴紧就会露缝。不是连续正常速度观看整节。初轮只验证两个向前上推；后续复核源的返回演示，再独立实测下述双向路线。

受测原创三页：33.87×19.05cm，同一浅灰底、开放custGeom路线；跨页连接点均x=24cm，前页y=19.05、后页y=0。三页路径各自可编辑，颜色C04457、3pt，页2/3为`push-up-1400`。路线到边缘裁切是有意行为，不能用自动缩进来“修复”越界而破坏接缝。

```powershell
python scripts/push_seam_build.py --officecli OFFICECLI --out NEW/route.pptx
python scripts/push_seam_verify.py NEW/route.pptx --report NEW/structure.json
powershell -NoProfile -File scripts/push_seam_export.ps1 -Deck NEW/route.pptx -Evidence NEW/native
powershell -NoProfile -File scripts/native_record_states.ps1 -Deck NEW/route.pptx -Plan assets/push-seam-record.json -Directory NEW/live -Ffmpeg FFMPEG
python scripts/push_seam_measure.py NEW/live/native.mp4 --native NEW/native --out NEW/measured --ffmpeg FFMPEG --ffprobe FFPROBE
```

OfficeCLI完成全部制作，PowerPoint只读导图/放映。构建与结构审计仅标准库，像素审计需Pillow/NumPy和现有FFmpeg/FFprobe。输出拒覆盖，保留原稿，不Quit共享PowerPoint。通用录屏器使用两次COM Next，只证明两个Push运动，不声称物理点击交互。

包内fresh：`validation/forward-push-seam/`，结构、三张原生PNG、约8秒真实放映、独立接缝量测通过。路线真实开放path、跨页端点、stroke、两个Push dir=u/1400ms均独立检查。217个编码帧按原始FFprobe PTS解码，每第三帧采样；14个中间态由前后原生PNG垂直平移配准，平均颜色误差最大约1.094/255，接缝±8px范围每行均有路线，中心偏差≤0.5px。原生页和跨页联系图已目视。

投影窗口为16:10、页为16:9；压缩会使纯背景边缘两行变暗，所以量测先由纯背景定位内部中心，再按已知16:9确定真实720px页高，不能将718px内部区域拉伸代替页。已在同一fresh片重算，未重新拍摄或改变路线。PNG模板配准用的是最终修正文字避让后的页面。

验证器还在一个中间帧的内存副本中清除7行接缝，正确检出7行缺口；只检验测量器，不修改PPT或原片，也不冒称实际作品发生过该失败。早期CLI构建标签数量索引错误、文字压线和初版内部区域量测记录保留，正式fresh已修正。

## 前进与返回

同一制作脚本无需改写：页2/3仍设向上Push，原生放映中用Right前进、Left返回，返回时PowerPoint反向播放原转场。不要为返回另加自动跳页、COM Goto或假截图；这些行为不能代替观众实际按键。页面路径、文字和连接点沿用上面的可编辑底稿。

已保存[三页示例](../../examples/push-return/push-return.pptx)及[44秒连续原生播放](../../examples/push-return/playback.mp4)。本机实际按键路线1→2→3→2→1→2→3→2→1，8次物理按键、0次点击、0额外输入，覆盖每条相邻边的两个方向各两次。复看24个运动画面及三页原生导图；1228个编码帧按原始PTS每第三帧取样，58个中间态分别按向上/向下的页对配准，平均RGB误差最大低于1.2/255，接缝17行无断口、中心偏差≤0.5px。内存副本清除7行的负对照检出7行缺口，不修改原片。旧结构验证器只证明路径与向上转场配置；双向结论来自本轮实际播放，不改变其原有验收范围。

复现时在所有幻灯片放映模式下，依次按Right、Right、Left、Left，再重复一轮；每次等待转场完成。检查往返运动中的路线，而非只核最终页码。换成别的导航入口或放映模式需另验。

范围：本固定三页/原生方向键双向Push/竖直同x接点。斜线接点、背景纹理拼接、不同页面尺寸、快速打断及自定义缓动未测；改变连接形状、导航方式、方向或底色后应重新验证，不把本例扩展为所有Push组合。
