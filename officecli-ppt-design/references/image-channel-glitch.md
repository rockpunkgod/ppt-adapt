# 真图像通道分离与窄条故障（C17）

backend：`external-render`。本配方的核心逐像素R/GB分离与条区复制是在Python标准库离线确定性处理，再用OfficeCLI原生picture嵌入。PowerPoint只读导图验收；不把此专门流程当所有图片编辑的默认方法，也不声称Office原生绝无等价实现。这里与自产GIF素材准备不同：被蒸馏的核心视觉机制本身在外部执行。

适用用户明确要求图像通道故障/局部错移时。先保留原创源RGBA图。以相同y、两个不同x偏移采样R与GB，使它们确实来自图像的不同源坐标，而不是复制两份彩色文字。透明边越界采样为零；合成alpha取两份的最大值，各分量按源alpha预乘、以合成alpha反除，避免半透明块错误变黑。当前受验样本为640×360、8位编码RGB数值（不作线性光/ICC变换），R右移10像素、GB左移10像素。混合公式针对本配方的两组通道重组，不宣称完全等价Adobe所有图层混合条件。

选区错移采用“复制选中行→移动副本→alpha over原图”，保留底层原有图像；不是整张图平移或切掉源行。三个半开行区间[90,101)、[166,174)、[247,261)，偏移+28/-36/+20。选区外字节必须保持相同。

第4页明确标为近似：只在y80..279、每9行前2行，以向左回看19像素的亮度最大样本延伸到右侧。它是原创有界方向拖线，不是Adobe Wind/Blast算法；源方向、预设核或强度参数未还原，不列精确Wind通过。未实现源最后的冷暖画笔Overlay调色，也不把它冒充通道分离。

复现（所有输出使用新目录）：

1. `python image_glitch_assets.py --outdir <fresh-assets>`：生成原创图、分离图、错移图、拖线近似、model.json。
2. `python image_glitch_build.py --officecli <officecli.exe> --assets <fresh-assets> --outdir <fresh-base>`：OfficeCLI建4页并嵌入原PNG字节。
3. `python image_glitch_audit.py --assets <fresh-assets> --deck <fresh-base>/image-glitch.pptx --ffmpeg <ffmpeg.exe> --report <fresh-audit.json>`：独立解码完整通道/alpha/区域及ZIP关系/嵌入字节检查。
4. `powershell -File image_glitch_export.ps1 -Deck <pptx> -Evidence <fresh-native>`：本机PowerPoint只读4PNG；仅关本次稿，不Quit共享应用。
5. `python image_glitch_analyze.py --assets <fresh-assets> --native <fresh-native> --ffmpeg <ffmpeg.exe> --report <fresh-pixels.json>`：真实PNG对照原像素及alpha背景融合，随后人工看彩边、局部条移与近似标签。

脚本只验证本原创诊断图/固定参数；真实照片、ICC色彩空间、高位深、任意通道组合和连续动画需要额外验证。PPT中图像可替换、定位和裁剪，通道/条区参数不作为PowerPoint可编辑层；更改参数需重算PNG并重新嵌入。需本机已有OfficeCLI/PowerPoint/FFmpeg，不自动安装软件。

源：[50P课程P25](https://www.bilibili.com/video/BV1oPHgeiEgb/?p=25)2584.2–2787.8秒。完整指定转写已读，Ego实际2618/2728/2744/2787帧确认通道混合选项、选区副本、水平条区与Overlay；这些分段观察不是全课连续观看。Wind口述明确但本次帧未捕获对话框，不从口述补出方向与内核。作者人物图片及课件不随包分发；仅脚本生成原创资产。

包内fresh证据在外部`E:/ppt skills/verification/image-glitch-forward/`，不随Skill分发。PPT SHA256 `7e9572dd508356a028a9a8f4c032ae08d99ef4232d9d912628bbb25733f7f241`；schema0错误，18离线像素/关系/字节项、16原生图项、4拒绝覆盖保护项通过。原生4PNG直接导出0.61秒，每页2646–3069稳定样本RGB RMS2.26–2.47/255（预设上限3），半透明块RMS0–0.94，透明角等于页底色16/28/42。通道公式、选区复制公式与alpha逐像素0差，选区外0差，三条内14347像素改变。实际目视彩边和条区错移；第4页近似标题亦可见。这些结果不验证Adobe Wind准确还原。
