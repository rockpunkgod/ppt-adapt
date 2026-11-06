# 汉字骨架、共笔与主题替笔（C11）

backend：`officecli-native`。OfficeCLI创建全部可编辑文本与闭合custGeom笔画；PowerPoint只读导图用于验收。适用原创标题/字标的少量人工设计文字，本配方只验证“山川”，不是任意汉字字体生成器。

先保留可编辑原文。按笔画走势人工建闭合轮廓，直段用节点连接，圆端/转角用三次贝塞尔；横竖粗细、重心和字怀由设计者判断。每笔单独命名，便于选择窗格编辑。判断两个部件是否可共用方向及位置后，只保留一个实体笔画，并在设计模型记录两个角色。主题替换需删除被替笔画，由新的闭合曲线路径承担其结构功能；原笔画上叠图标不算替换。

本样本第1页原文“山川”采用本机Microsoft YaHei Bold；第2页7笔手画骨架；第3页山右竖/川左竖共用shared-stroke，6对象；第4页移除chuan-middle，改用river-symbol河流曲线，仍6对象。全部自定义几何由标准库编排OfficeCLI batch raw-set，不依赖研究PPT、字体转曲或图片。

制作：`python chinese_stroke_build.py --officecli <officecli.exe> --outdir <fresh-dir>`。
原生验收：`powershell -File chinese_stroke_export.ps1 -Deck <fresh-dir>/chinese-strokes.pptx -Evidence <fresh-native-dir>`；只关闭本次打开的文稿，不退出共享PowerPoint。
结构：`python chinese_stroke_verify.py --deck <pptx> --model <design-model.json> --report <fresh-report.json>`。
像素：`python chinese_stroke_analyze.py --native <native-dir> --ffmpeg <ffmpeg.exe> --report <fresh-report.json>`。
所有输出必须新路径，拒绝覆盖。需本机已有PowerPoint、OfficeCLI和FFmpeg；不自动安装字体或软件。

验收不能只数节点：检查原文字备份、共笔无重复对象、原中竖确实删除、三次曲线和闭合路径；原生图应看到独立骨架4连通块、共笔/替换各3块，其他5笔保持完全相同，替换区显著变化。最后人工看“山川”是否可辨、河流是否适合主题。连通块与像素指标只作几何证据，不等于语言识读能力；未经盲测和小字号测试。

源：[50P课程](https://www.bilibili.com/video/BV1oPHgeiEgb/) P20 3521.9–3836.5、P22 3464.5–3891.5、P26 348.6–1094.4、P27 456.4–912.6、P48 1074–1751秒。已读完整指定转写/每课6分布帧，另实际取P26 1057/P27 901帧；不是全课程连续观看。将低透明字体参照→人工骨架→节点/圆角→共笔删除重复→主题替换归为一个工作流。商业“左洱”“夏天”等源外观不捆绑、不宣称本样本等价复刻。源识读优先保留，建筑受力类比不作为字体普遍定律。

不覆盖：任意中文字符集、自动拓扑推断、活字体/字库导出、Adobe路径算法等价、自动审美评分、源所有视觉变体。形状编辑会改变标题图形，原文备份保持；变更后必须重新看原生图和识读。

包内fresh验证：外部`E:/ppt skills/verification/chinese-stroke-forward/`。重建PPT SHA256 `daad3b9632e3e51ed9407cf8efc61fb3b0000fb6b43f31b79cea8ab7b78f6a42`；OfficeCLI schema0错误，17结构项、5原生像素项及2拒绝覆盖项通过。原生只读4PNG导出0.53秒；共笔/替笔图已目视。未改笔画区RMS0，替换区域RMS72.63/255。证据和PPT留在源包外，不作为技能资产分发；没有动画/物理交互通过声明。
