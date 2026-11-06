# 从零构造双面3D卡片（C19）

来源：大师课9Ewb0Rd3P3E，3:36:50–3:49:30。使用场景：真正需要翻面显示不同内容，且接受卡片两面内容作为调用方PNG纹理。后端officecli-native：本地准备GLB资产、OfficeCLI原生3D对象/Morph；PowerPoint仅用于原生验收。不能称为外部3D视频或2D图片翻转。

操作顺序：

1. 准备同尺寸9:16 RGBA PNG两面，文字方向均正常，不预先镜像。调用方文件仅在本机读；两面均以原字节嵌入GLB。没有素材时，使用脚本自产F/B诊断fixture。
2. 本地GLB包含两个矩形三角面，法线分别±Z，背面索引绕序反转、UV的U方向反转；两个材质都`doubleSided=false`。`alphaMode=MASK`保留PNG透明圆角。平面相距0.002模型单位，不承诺实体厚边。
3. OfficeCLI `model3d`的`src`指向该GLB，三页保持名称`!!dual-card`，真实`rotation=0,0,0`、`0,180,0`、`0,0,0`，页转场`morph-2000`。制作结束先close写盘，再交PowerPoint。
4. 结构审计不能只数GLB：解码accessor，检查绕序叉积/法线，材质单面、双纹理字节、UV、PPT每页真model3d和旋转值。
5. 本机PowerPoint确认对象Type30、两面文字方向正确，再看翻转中间态。需有真透视梯形和接近侧缘相位；正反面不同时混叠。导出电影与真实桌面放映证据分开记录。

输入边界：调用方照片/排版已经完成；不是自动生成背面排版工具。只测PNG/单个矩形卡片/静态GLB/原生Morph，不外推任意GLB、骨骼动画、物理光照材质或复杂透明混合。圆角为纹理alpha裁切，不是圆角网格。PowerPoint灯光会改变显示颜色。

依赖：OfficeCLI1.0.160，支持3D和Morph的本机PowerPoint；Python Pillow仅供本地PNG检查与诊断纹理生成，独立像素测量另需NumPy，ffmpeg用于录屏及取帧。没有将用户图片上传到建模网站的步骤。


复现入口（输出目录必须不存在；以下为包内脚本，不依赖研究目录成品）：

```powershell
& $Python $Skill/scripts/glb_card_build.py --officecli $OfficeCLI --out $Out
& $Python $Skill/scripts/glb_card_verify.py $Out
powershell -NoProfile -ExecutionPolicy Bypass -File $Skill/scripts/glb_card_native.ps1 -Deck "$Out/card.pptx" -Evidence "$Out/native"
& $Ffmpeg -hide_banner -loglevel error -i "$Out/native/native.mp4" -vf fps=6 "$Out/native/f-%03d.png"
& $Python $Skill/scripts/glb_card_measure.py "$Out/native"
powershell -NoProfile -ExecutionPolicy Bypass -File $Skill/scripts/glb_card_record.ps1 -Deck "$Out/card.pptx" -Evidence "$Out/live" -Ffmpeg $Ffmpeg
& $Ffmpeg -hide_banner -loglevel error -i "$Out/live/card-live.mp4" -vf fps=12 "$Out/live/f-%03d.png"
& $Python $Skill/scripts/glb_card_measure.py "$Out/live" --fps 12
```

`$Skill`指本Skill目录，`$Python`为含Pillow/NumPy的本机Python；`$OfficeCLI`/`$Ffmpeg`指可执行文件。含空格时给整个参数加引号；原生UI仅能独占一份放映测试。原生导片和实际桌面片分别约9秒和8秒。脚本只关闭自己的PPT，不Quit，不操作ChartData。

调用方素材：在build入口加`--front $FrontPng --back $BackPng`，两张同时给，PNG必须RGBA、同尺寸、9:16。PNG原字节写入GLB，不修改图像、不外传。`verify`检查二进制网格/纹理/OPC；`measure`颜色阈值只适用于默认自产F/B诊断fixture，任意用户图片应另定语义位置与人工原生对照，不能把默认fixture的颜色计数规则外推。

验收记录应分别列出：GLB头/对齐/法线绕序/UV/双材质、PPT实际Type30与旋转、正反向文字、透明角、两轮翻面透视/侧缘、两次真实点击到2/3页。schema或兼容占位PNG不能代证原生3D效果。通过范围不包含GLB骨骼/动画/厚实体。

包内exact前向：`validation/forward-glb-card/README.md`。全新模型/两张纹理/3页PPT由包内builder自产；结构/schema/拒覆盖、原生Type30与0/180/0旋转、55帧原生导片和96帧实际桌面全目视、两次物理点击到2/3页、两套独立运动像素审计、原稿SHA保持均通过。录屏宽314px→侧缘10/20px，中间左右边高度差75–148px；真模型结构和透视共同排除2D压宽换图。此为默认诊断fixture通过，不把其颜色阈值外推给调用方任意纹理。
