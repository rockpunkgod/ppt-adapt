# 虚实之间

`glyph-gaussian.pptx`用原创四笔“木”展示三种画面：全字退后、只虚化一笔、深色背景下的虚实组合。原始轮廓、清晰笔画和原文字均保留；虚化部分为透明PNG，可通过脚本重建。

制作复用已有笔画路径，由OfficeCLI建立页面和清晰形状；外部计算对选中图层的alpha执行显式Gaussian卷积。它属于external-render，不是PowerPoint原生艺术Blur或方向性Motion Blur。生成器保留alpha源、参数与图层，不需要下载素材。

运行`glyph_gaussian_build.py --officecli EXE --outdir NEW --sigma 12`可复现，需本机已有Pillow/NumPy环境。详见[制作指导](../../officecli-ppt-design/references/local-blur-background.md)。当前验证限原创四笔与默认σ12；新字体、其他参数或自动拆笔不能直接沿用通过结论。

`preview.png`为第2页原生导图，`playback.mp4`为23秒本机PowerPoint实播，覆盖深浅底色、返回和再次进入。三页保存重开画面一致；此样稿为静态视觉，没有动画。
