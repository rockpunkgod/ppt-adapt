# 扫描笔触：反相滤色与限定真alpha（C16）

这是两个能力，不能互相代称。源教程实际操作为白纸手写/拍照或扫描→反相→滤色Screen视觉去底；独立真alpha提取是本配方的额外改进。素材均由程序绘制原创曲线/细线并模拟扫描，明确不是实物扫描、字体描边、InkML或作者字迹。

**A. 源机制：固定底色反相Screen，`officecli-native`。** 保留中性白纸诊断PNG原字节。OfficeCLI给blip附a:duotone白→黑实现中性灰阶反相，随后a:fillOverlay blend=screen，solidFill对应固定背景。公式为1-(1-I)*(1-B)。图像实际仍不透明，是固定背景的视觉合成，不是真alpha，不会自动采样下面的复杂图层；背景改变必须重设effect颜色并原生复验。只验证中性白纸/灰阶笔触及三种固定solid背景，不等价所有Adobe混合/彩色照片。

本机量化校准：未补偿v1原生纸区比背景各通道低1，严格纸区RMS约1.02，超过预设.5，因此false保留。v2仅将Screen输入各通道+1，页底保持不变：实际页底162A38→effect输入172B39，D7AD59→D8AE5A，718E79→728F7A。只列这三组实测校准，不能把+1当所有颜色/其他渲染器/动态背景的规律。阈值仍为纸RMS≤.5、整体公式RMS≤2；没有放宽。

**B. 独立改进：均匀纸色/单墨色真alpha，`external-render`。** 用户明确要求提取此类图像且能给出均匀纸色/墨色时，才运行确定性提取。保留原PNG和参数，用颜色轴投影估计覆盖率：alpha=clamp(round(255*dot(paper-pixel,paper-ink)/dot(paper-ink,paper-ink)))。当前诊断paper248/246/239、ink24/30/26；输出白色笔触RGBA，没有二值阈值/消噪/模糊，保灰边、淡线及细尖端。算术为8位编码RGB，不作线性光/ICC推论。输出alpha是真图像透明度，换三种背景无需改mask。

只验证均匀、已校准纸色、单墨色、无阴影/纸纹的原创模拟图，不是任意扫描/摄影抠图器。不会自动判断真实照片的纸色，既有alpha输入会拒绝；需要额外去阴影/纸纹时不擅自阈值删细线。核心alpha提取发生在外部，OfficeCLI只做基础PPT与嵌入原字节，不能标为原生抠图。一般图片编辑不默认走此专门工程。

复现，全部使用新输出路径：

1. `python scan_stroke_fixture.py --outdir <fixture>` 生成明确标注的模拟扫描、中性原图、真值alpha和模型。
2. A：`python scan_stroke_screen_build.py --officecli <exe> --fixture <fixture> --screen-byte-compensation 1 --outdir <screen-base>`。
3. B：`python scan_stroke_extract.py --input <fixture>/simulated-scan.png --paper 248,246,239 --ink 24,30,26 --color FFFFFF --ffmpeg <exe> --outdir <extraction>`；`python scan_stroke_build.py --officecli <exe> --extraction <extraction> --outdir <alpha-base>`。
4. 分别`powershell -File scan_stroke_export.ps1 -Deck <pptx> -Evidence <new-native-dir>`，只读4PNG并关闭本次稿，不Quit共享PowerPoint。
5. A：`python scan_stroke_screen_verify.py --fixture <fixture> --deck <screen-pptx> --native <screen-native> --screen-byte-compensation 1 --ffmpeg <exe> --report <new-report>`。
6. B：`python scan_stroke_verify.py --fixture <fixture> --extraction <extraction> --deck <alpha-pptx> --ffmpeg <exe> --report <new-report>`；`python scan_stroke_analyze.py --extraction <extraction> --native <alpha-native> --ffmpeg <exe> --report <new-report>`。

需本机已有OfficeCLI/PowerPoint/FFmpeg；不自动安装。实际调用方PNG只限用户明确要求的已知纸/墨提取流程，保留原件并显式传校准参数；诊断真值验收脚本用于本自产fixture，不能凭此宣称新真实照片通过。

源：[50P P20](https://www.bilibili.com/video/BV1oPHgeiEgb/?p=20)1745.9–1834.5、[P25](https://www.bilibili.com/video/BV1oPHgeiEgb/?p=25)2181–2403.3秒指定转写完整读取；Ego精确1803/1812/1826及P25 2212整帧已看。1812菜单及1826面板明确“滤色”，转写“绿色”为误识。P25多为风格/输入方式，未提供自动抠图算法。第三方网站“无版权随便用”的口述不作为许可，未复制其资产。

阴性保留：合法a:duotone alpha端点100000/0的本机探针schema过，但实际两页整块白，纸RGB255/255/255，未得到透明；不能因此宣称原生真alpha提取通过或所有PowerPoint原生方法都不可能。此阴性与源Screen合成无关。

包内fresh证据在Skill外`E:/ppt skills/verification/scan-stroke-forward/`，两稿原生8PNG0.91秒依次导出。Screen18项、alpha15离线+9原生项、5覆盖保护项全过；两稿schema0。Screen纸区RMS.208/.117/.143（阈.5），公式.308/.171/.208（阈2）；真alpha纸残留RMS0、每底158个12/255淡线点保留，原真值覆盖率最大误差1、252级透明度、4959灰边像素及216173透明纸像素。实际图已目视；没有放宽阈值。Screen稿SHA256 `ff3b7dd447019dea806c1078ca2b34d6bdd7ee3831eb6a5aceb1649560dae93d`，alpha稿 `413d55957deb322157ae4b881c943c34364944b941b76f30511bf6993010c019`。证据/测试稿不随Skill打包。8位输入保护增加后重新前向提取，输出alpha字节hash仍为`90dfb307a7b2012c7a93f98b53e32b5b78abb9bac2d0e4e5d407ae15c3fdd4fd`。
