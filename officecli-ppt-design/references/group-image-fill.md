# 多片统一图片坐标：真实父组填充

适用：将一幅图分为相邻或分离的多个原生窗口，同时保持原图坐标关系。backend **officecli-native**；PowerPoint COM只读检查与导图。图片由调用方提供，库不捆教程图片或演示照片。

机制：OfficeCLI建立4个矩形与真实group；先在一个子形状放图以建立关系，再经OfficeCLI `raw-set` 向父 `p:grpSpPr` 加 `a:blipFill`（同一图片rId、stretch），将4个子形状的直接fill替为 `a:grpFill`。父组的xfrm/off/ext/chOff/chExt决定统一图片坐标。子形状直接图片填充必须移除，否则会遮盖继承。`Group.Fill.UserPicture` 在本机实测会逐片重复整图，不能替代该结构。

## 可复现入口

下列参数分别是当前Python、OfficeCLI、调用方JPEG、尚不存在的输出路径；JPEG暂支持普通/渐进SOF0/1/2。不裁切原图，按其宽高比放入限宽27.5cm/限高10.7cm区域。依赖：构建/结构核验仅标准库；原生导图Windows PowerPoint；像素核验另需已有Pillow运行环境（不自动安装）。

```powershell
& $Py "$Skill/scripts/bili_layout_group_image_build.py" --officecli $OfficeCLI --image $Image --out "$Out/fragments.pptx"
& $Py "$Skill/scripts/bili_layout_group_image_verify.py" "$Out/fragments.pptx" --image $Image --report "$Out/structure.json"
& "$Skill/scripts/bili_layout_group_image_export.ps1" -Deck "$Out/fragments.pptx" -Evidence "$Out/native"
& $PillowPy "$Skill/scripts/bili_layout_group_image_pixels.py" "$Out/native" --structure "$Out/structure.json" --report "$Out/pixels.json"
```

输出3页：完整原图基准、2×2相邻四片、四角分离窗口。构建完成显式关闭OfficeCLI驻留文稿，避免PowerPoint读到未落盘版本。已有文稿/伴随证据/导图目录拒绝覆盖。导图只关闭自身文稿，不Quit共享PowerPoint。

验收：CLI schema；图片实际SHA256一致；真实group×2，各4子形状；父组唯一图片fill、子形状grpFill且无直接图片fill；独立组坐标变换与预期窗口位置；COM Type6/GroupItems4；8片内部区域逐像素对比第1页同坐标，排除边缘5px、RGB RMS<0.2/255，人工复看拼图与分离窗口。

## 证据与边界

源 BV1VTDGYsEmd P06，945–1074s、1914–1947s OCR与1035s/1938s整帧：先组合再填图、各块呈同一图片不同部分。源为旋转菱形/椭圆切片；本配方原创矩形窗口只验证共同图片坐标机制，未宣称复原任意旋转/自由曲线。

独立前向 `validation/forward-group-image/`：3页原生PNG复看通过，8块RMS0.014–0.045，阈值0.2；结构与拒覆盖通过。研究失败探针保留在 `distill/bilibili-layout-type/course-27/image-fragments-v1/`：COM UserPicture重复整图，RMS约73–92；逐片srcRect几何正确但单独重采样RMS3.3–4.6，未凭放宽阈值升级，最终采用真实父组继承。

范围：轴对齐矩形、当前整组共同图片坐标。未测随意子形状移动/旋转后的重新映射、Morph或动态继承。改变窗口/父组边界后重跑结构、原生及像素核验；不能拿本样例代证所有几何。
