# 局部大风拖尾与双层冷暖叠加

先完成清晰内容，再把故障视觉限定在几条细带。保留底图，条带只做局部断裂；冷暖光从两个方向组织画面，不遮住标题。核心图像处理归 **external-render**，OfficeCLI负责页面、原生文字和独立图片对象。

## 制作

复用 `image_glitch_assets.py` 的原创640×360图像、R/GB错位及三条带位置，不改旧 `streak-approx.png` 的近似语义。`image_blast_assets.py`只准备透明条带和两张柔和画笔素材，不实现或冒充Wind。

```powershell
python scripts/image_blast_assets.py --channels ORIGINAL/channels.png --outdir NEW_STRIPS
# 在图像编辑器打开NEW_STRIPS/strip-copy.png；滤镜→风格化→风，
# Blast / From the Right，确定；保留640×360和透明度导出ACTUAL_BLAST.png。
python scripts/image_blast_assets.py --channels ORIGINAL/channels.png --blast ACTUAL_BLAST.png --outdir NEW_ASSETS
```

在编辑器打开`NEW_ASSETS/blast-base.png`；分别导入`magenta-brush.png`、`cyan-brush.png`为独立图层，保持整幅画布对齐。底图→洋红→青色，每个画笔层分别设 **Overlay / 叠加，图层不透明度100%**。画笔PNG本身带柔和alpha，100%图层不透明度不表示每个像素全不透明。导出`NEW_ASSETS/overlay-two-photopea.png`。

不要先合并两色再Overlay；Overlay依赖其下方颜色，组合顺序会改变结果。本例实际两种做法有215783个不同像素。需要改光色时重做两层合成，不能用PPT图片透明度模拟Overlay。

```powershell
python scripts/image_blast_build.py --officecli EXE --assets NEW_ASSETS --out NEW_DECK.pptx
```

脚本沿用现有Deck机制制作四页：通道、实际条带、双层冷暖、原创标题应用。第2页底图与Blast条带是两个独立图片，第3/4页为外部合成PNG；标题、位置和图片可编辑，滤镜没有变成PowerPoint原生效果。适配新照片时重新选条带、位移和光斑；当前准备器明确限640×360原创诊断图，不按固定坐标生套任意素材。

## 已验证的具体组合

[四页PPTX、素材和27秒原生播放](../../examples/image-blast/README.md)。实际Photopea风滤镜选择Blast与From the Right后导出透明条带；另实际执行Blast反方向及普通Wind作对照。原始条带与实际Blast有7850个变化像素、289个新增alpha像素，变化只在33行选区；反方向与本方向有10307个不同像素，普通Wind与Blast有8311个不同像素。各选项确实改变输出，不能仅凭菜单点击声称已实现。

两个独立画笔层均实际设置Overlay100%后导出。按已有encoded-sRGB Overlay定义计算的参考与实际编辑器输出最大通道差3、平均0.222；保留实际导出，不宣称像素等同Adobe Photoshop或其ICC/gamma设置。

最终本机PowerPoint窗口放映27秒，6次物理方向键走1→2→3→4→3→2→3，无额外输入；7个分布实播状态复看。为避开桌面浮层，只移动本次放映窗口再录，未关闭其他程序；早期被遮挡的录像未交付。四页保存重开原生导图逐像素相同，素材字节嵌入与文字无溢出已核。窗口放映和高分辨率导出存在文字/边缘重栅格化差异，不声称两者像素一致。

该结果覆盖静态图像组合与上述前进返回路线，不覆盖动态故障闪烁、任意照片、PPT原生混合模式或快速打断。

来源操作依据：BV1oPHgeiEgb P25，2743–2751秒为Wind→Blast、从右及条带拖尾，2755–2788秒为两张空层、洋红/青色软笔与Overlay100%。实际课程画面已复核，原图和字体不随技能包分发；这里的几何图像与排版为原创。
